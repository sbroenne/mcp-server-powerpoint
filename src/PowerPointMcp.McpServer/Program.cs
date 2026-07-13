using System.IO.Pipelines;
using System.Runtime.InteropServices;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.McpServer;

/// <summary>
/// PowerPointMCP Model Context Protocol (MCP) server host.
/// Exposes Core presentation commands as MCP tools over stdio (production) or an in-memory
/// pipe transport (tests). All logging is routed to stderr — stdout is reserved for JSON-RPC.
/// </summary>
public class Program
{
    private static readonly object TestTransportLock = new();

    // Test transport configuration — set by tests (in-memory pipes) before calling Main().
    // Intentionally static for test injection, but guarded so leaked test state fails fast
    // instead of contaminating the next transport-backed test.
    private static Pipe? _testInputPipe;
    private static Pipe? _testOutputPipe;
    private static CancellationTokenSource? _testShutdownCts;
    private static long _testTransportGeneration;

    /// <summary>
    /// Configures the server to use an in-memory pipe transport for testing.
    /// Call before running <see cref="Main"/> to enable test mode.
    /// </summary>
    /// <param name="inputPipe">Pipe the client writes requests to and the server reads.</param>
    /// <param name="outputPipe">Pipe the server writes responses to and the client reads.</param>
    public static void ConfigureTestTransport(Pipe inputPipe, Pipe outputPipe)
    {
        lock (TestTransportLock)
        {
            if (_testInputPipe != null || _testOutputPipe != null || _testShutdownCts != null)
            {
                throw new InvalidOperationException(
                    "Test transport is already configured. Ensure the previous MCP transport test completed cleanup before starting another one.");
            }

            _testInputPipe = inputPipe;
            _testOutputPipe = outputPipe;
            _testShutdownCts = new CancellationTokenSource();
            _testTransportGeneration++;
        }
    }

    /// <summary>
    /// Requests shutdown for the active in-memory test transport without clearing transport state.
    /// </summary>
    public static void RequestTestTransportShutdown()
    {
        CancellationTokenSource? shutdownCts;

        lock (TestTransportLock)
        {
            shutdownCts = _testShutdownCts;
        }

        if (shutdownCts != null)
        {
            try
            {
                shutdownCts.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // ResetTestTransport() owns disposing the test CTS after the host has fully stopped.
            }
        }
    }

    /// <summary>
    /// Resets test transport configuration after the in-memory test host has stopped.
    /// </summary>
    public static void ResetTestTransport()
    {
        CancellationTokenSource? shutdownCts;

        lock (TestTransportLock)
        {
            shutdownCts = _testShutdownCts;
            _testShutdownCts = null;
            _testInputPipe = null;
            _testOutputPipe = null;
        }

        shutdownCts?.Dispose();
    }

    /// <summary>
    /// Entry point. Starts the MCP host and blocks until shutdown, then guarantees all PowerPoint
    /// sessions are disposed so no POWERPNT.exe process lingers.
    /// </summary>
    public static async Task<int> Main(string[] args)
    {
        // Handle --help and --version for easy verification without launching the stdio server.
        if (args.Length > 0)
        {
            var arg = args[0].ToLowerInvariant();
            if (arg is "-h" or "--help" or "-?" or "/?" or "/h")
            {
                ShowHelp();
                return 0;
            }

            if (arg is "-v" or "--version")
            {
                ShowVersion();
                return 0;
            }
        }

        Pipe? testInputPipe;
        Pipe? testOutputPipe;
        CancellationTokenSource? testShutdownCts;
        long testTransportGeneration;

        lock (TestTransportLock)
        {
            testInputPipe = _testInputPipe;
            testOutputPipe = _testOutputPipe;
            testShutdownCts = _testShutdownCts;
            testTransportGeneration = testInputPipe != null && testOutputPipe != null
                ? _testTransportGeneration
                : 0;
        }

        var builder = Host.CreateApplicationBuilder(args);

        ConfigureStdioLogging(builder.Logging);

        // In-process PowerPointMcpService: the SAME Service class the CLI daemon hosts over a
        // named pipe, but here consumed directly, in-process, with no pipe — mirroring
        // mcp-server-excel's ServiceBridge architecture (one shared Service class, two hosting
        // modes). MCP tools never call the service's string-command dispatch; they only need its
        // session registry, so that's the only thing exposed to DI. One long-lived PowerPoint
        // session per id across many tool invocations; disposed on host shutdown by
        // PresentationSessionShutdownService.
        builder.Services.AddSingleton<PowerPointMcpService>();
        builder.Services.AddSingleton(sp => sp.GetRequiredService<PowerPointMcpService>().Sessions);
        builder.Services.AddHostedService<PresentationSessionShutdownService>();

        var mcpBuilder = builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "powerpoint-mcp",
                    Version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "0.1.0"
                };

                options.ServerInstructions = """
                    PowerPointMCP automates Microsoft PowerPoint via COM interop (Windows only).

                    SESSION LIFECYCLE (all via the single "presentation" tool's action parameter):
                    1. presentation(action=create, filePath) — create a new deck and open it; returns a sessionId.
                    2. presentation(action=open, filePath) — open an existing deck; returns a sessionId.
                    3. Pass that sessionId to all subsequent tools.
                    4. presentation(action=save, sessionId) — persist changes.
                    5. presentation(action=close, sessionId) — release the PowerPoint process when done.

                    Use presentation(action=list) to see which sessions are currently open.
                    Always provide full Windows paths (e.g. C:\\Users\\me\\Documents\\deck.pptx).
                    """;
            })
            .WithToolsFromAssembly();

        if (testInputPipe != null && testOutputPipe != null)
        {
            // Test mode: in-memory pipe transport so tests can drive the server without stdio.
            mcpBuilder.WithStreamServerTransport(
                testInputPipe.Reader.AsStream(),
                testOutputPipe.Writer.AsStream());
        }
        else
        {
            // Production mode: stdio transport (stdout = JSON-RPC channel).
            mcpBuilder.WithStdioServerTransport();
        }

        var host = builder.Build();

        // Capture the singleton reference now — after RunAsync returns the host has disposed its
        // service provider, so resolving it in the finally would throw ObjectDisposedException.
        var service = host.Services.GetRequiredService<PowerPointMcpService>();

        var runToken = testShutdownCts?.Token ?? CancellationToken.None;

        var stdinMonitor = testInputPipe == null
            ? StdinPipeMonitor.Start(host.Services.GetRequiredService<IHostApplicationLifetime>())
            : null;

        try
        {
            await host.RunAsync(runToken);
            return 0;
        }
        catch (OperationCanceledException)
        {
            // Graceful shutdown via cancellation (Ctrl+C, test shutdown) — expected, not an error.
            return 0;
        }
        finally
        {
            stdinMonitor?.Dispose();

            // Backstop: guarantee no lingering POWERPNT.exe even if the hosted service did not run.
            // Dispose is idempotent, so double-disposal on normal shutdown is safe.
            service.Dispose();
        }
    }

    /// <summary>
    /// Routes all console logging to stderr. MCP stdio reserves stdout exclusively for JSON-RPC
    /// frames, so any log written to stdout would corrupt the protocol channel.
    /// </summary>
    internal static void ConfigureStdioLogging(ILoggingBuilder logging)
    {
        logging.ClearProviders();
        logging.AddConsole(consoleLogOptions =>
        {
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
        });
        logging.SetMinimumLevel(LogLevel.Warning);
    }

    private static void ShowHelp()
    {
        var version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "0.1.0";
        Console.WriteLine($"""
            PowerPoint MCP Server v{version}

            An MCP (Model Context Protocol) server for Microsoft PowerPoint automation.

            Usage:
              Sbroenne.PowerPointMcp.McpServer.exe [options]

            Options:
              -h, --help      Show this help message
              -v, --version   Show version information

            Without options, starts the MCP server in stdio mode.

            Requirements:
              - Windows x64
              - Microsoft PowerPoint (desktop) installed
            """);
    }

    private static void ShowVersion()
    {
        var version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "0.1.0";
        Console.WriteLine($"PowerPoint MCP Server v{version}");
    }
}

/// <summary>
/// Hosted service whose sole job is to dispose the in-process <see cref="PowerPointMcpService"/>
/// (and thereby every open presentation session) when the host stops — closing each PowerPoint
/// process so none is orphaned after MCP client disconnect, Ctrl+C, or test shutdown.
/// </summary>
internal sealed class PresentationSessionShutdownService(PowerPointMcpService service) : IHostedService
{
    public Task StartAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    public Task StopAsync(CancellationToken cancellationToken)
    {
        service.Dispose();
        return Task.CompletedTask;
    }
}

/// <summary>
/// Monitors the stdin pipe to detect when the parent MCP client process exits.
/// </summary>
/// <remarks>
/// MCP stdio transport keeps the server alive via stdin/stdout pipes. A clean parent exit closes
/// the pipe and the SDK shuts the host down. But if the parent is killed (Task Manager, SIGKILL),
/// the SDK may not notice. This monitor polls the stdin handle to detect the broken pipe and
/// trigger graceful shutdown, ensuring COM handles are released and PowerPoint is not orphaned.
/// Only activates when stdin is a named pipe (the normal stdio MCP transport case); terminal,
/// debugger, and file-redirection launches are left to their own shutdown mechanisms.
/// </remarks>
internal static class StdinPipeMonitor
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool PeekNamedPipe(
        IntPtr hNamedPipe, IntPtr lpBuffer, uint nBufferSize,
        IntPtr lpBytesRead, out uint lpTotalBytesAvail,
        out uint lpBytesLeftThisMessage);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll")]
    private static extern uint GetFileType(IntPtr hFile);

    private const int StdInputHandle = -10;
    private const uint FileTypePipe = 0x0003;
    internal const int ErrorBrokenPipe = 109;   // ERROR_BROKEN_PIPE
    internal const int ErrorNoData = 232;       // ERROR_NO_DATA (write end closed)

    /// <summary>
    /// Starts the stdin pipe monitor, or returns <see langword="null"/> when stdin is not a pipe
    /// (terminal, debugger, file redirection) — those cases don't need broken-pipe detection.
    /// </summary>
    public static Timer? Start(IHostApplicationLifetime lifetime)
    {
        var handle = GetStdHandle(StdInputHandle);
        if (handle == IntPtr.Zero || handle == new IntPtr(-1))
            return null;

        if (GetFileType(handle) != FileTypePipe)
            return null;

        return StartCore(lifetime, handle);
    }

    internal static Timer StartCore(IHostApplicationLifetime lifetime, IntPtr pipeHandle,
        TimeSpan? pollInterval = null)
    {
        var interval = pollInterval ?? TimeSpan.FromSeconds(5);
        return new Timer(_ =>
        {
            if (!PeekNamedPipe(pipeHandle, IntPtr.Zero, 0, IntPtr.Zero, out uint _, out uint _))
            {
                var error = Marshal.GetLastWin32Error();
                if (error is ErrorBrokenPipe or ErrorNoData)
                    lifetime.StopApplication();
            }
        }, null, interval, interval);
    }
}
