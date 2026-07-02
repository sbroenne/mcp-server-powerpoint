// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.PowerPointMcp.ComInterop;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Drives the real <see cref="Program"/> MCP host over an in-memory pipe transport so tests can
/// exercise the exact production code path (DI, tool discovery, JSON-RPC serialization) without
/// stdio or an external process. Ported from mcp-server-excel's
/// <c>ProgramTransportTestHost</c> — only the shutdown timeout source (PowerPointMcp's
/// <see cref="ComInteropConstants"/>) differs.
/// </summary>
internal static class ProgramTransportTestHost
{
    private static readonly TimeSpan ClientInitializationTimeout = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan ServerReadyTimeout = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan ServerReadyRetryDelay = TimeSpan.FromMilliseconds(50);
    // Must exceed PresentationSessionRegistry's own internal DisposeAll timeout
    // (StaThreadJoinTimeout + 30s — see PresentationSessionRegistry.DisposeAllTimeout), or this
    // outer wait could time out and force-cancel the host WHILE DisposeAll is still legitimately
    // draining a slow-quitting PowerPoint session, turning a benign slow shutdown into a false
    // test failure. +60s gives 30s of headroom over that inner bound.
    private static readonly TimeSpan ServerShutdownTimeout =
        ComInteropConstants.StaThreadJoinTimeout + TimeSpan.FromSeconds(60);

    /// <summary>
    /// Configures the in-memory transport, starts the real server host (<c>Program.Main</c>) on a
    /// background task, and connects an MCP client over the same pipes.
    /// </summary>
    public static async Task<(McpClient Client, Task ServerTask)> StartAsync(
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        string clientName,
        CancellationToken cancellationToken)
    {
        Program.ConfigureTestTransport(clientToServerPipe, serverToClientPipe);

        var serverTask = Program.Main([]);
        var client = await ConnectClientWithRetryAsync(clientToServerPipe, serverToClientPipe, clientName, cancellationToken);

        return (client, serverTask);
    }

    /// <summary>
    /// Tears everything down in order: dispose the client, request test-transport shutdown,
    /// complete the pipes so the host's stream transport observes EOF, await the server task, and
    /// finally reset the static test-transport state so the next test can configure it again.
    /// </summary>
    public static async Task StopAsync(
        McpClient? client,
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        Task? serverTask,
        ITestOutputHelper output,
        CancellationTokenSource? cancellationTokenSource = null)
    {
        if (client != null)
        {
            try
            {
                await client.DisposeAsync();
            }
            catch (Exception ex)
            {
                output.WriteLine($"Warning: Failed to dispose MCP client cleanly: {ex.Message}");
            }
        }

        if (serverTask == null)
        {
            await TryCompleteAsync(clientToServerPipe.Writer, output, nameof(clientToServerPipe) + ".Writer");
            await TryCompleteAsync(serverToClientPipe.Reader, output, nameof(serverToClientPipe) + ".Reader");
            await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
            await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");
            Program.ResetTestTransport();
            return;
        }

        Program.RequestTestTransportShutdown();
        await TryCompleteAsync(clientToServerPipe.Writer, output, nameof(clientToServerPipe) + ".Writer");
        await TryCompleteAsync(serverToClientPipe.Reader, output, nameof(serverToClientPipe) + ".Reader");

        try
        {
            await serverTask.WaitAsync(ServerShutdownTimeout);
        }
        catch (OperationCanceledException)
        {
        }
        catch (TimeoutException)
        {
            output.WriteLine("Warning: MCP test host did not stop within timeout; forcing cancellation.");

            if (cancellationTokenSource is not null && !cancellationTokenSource.IsCancellationRequested)
            {
                await cancellationTokenSource.CancelAsync();
            }

            try
            {
                await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
                await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");
                await serverTask.WaitAsync(ServerShutdownTimeout);
            }
            catch (OperationCanceledException)
            {
            }
            catch (TimeoutException)
            {
                output.WriteLine("Warning: MCP test host still did not stop after forced cancellation.");
            }
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: MCP test host faulted during shutdown: {ex.Message}");
        }

        await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
        await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");

        if (!serverTask.IsCompleted)
        {
            throw new TimeoutException("MCP test host did not stop after shutdown, forced cancellation, and pipe completion.");
        }

        Program.ResetTestTransport();
    }

    private static async Task TryCompleteAsync(PipeWriter writer, ITestOutputHelper output, string pipeName)
    {
        try
        {
            await writer.CompleteAsync();
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: Failed to complete {pipeName}: {ex.Message}");
        }
    }

    private static async Task TryCompleteAsync(PipeReader reader, ITestOutputHelper output, string pipeName)
    {
        try
        {
            await reader.CompleteAsync();
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: Failed to complete {pipeName}: {ex.Message}");
        }
    }

    private static async Task<McpClient> ConnectClientWithRetryAsync(
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        string clientName,
        CancellationToken cancellationToken)
    {
        var deadline = DateTime.UtcNow + ServerReadyTimeout;

        while (true)
        {
            try
            {
                return await McpClient.CreateAsync(
                    new StreamClientTransport(
                        serverInput: clientToServerPipe.Writer.AsStream(),
                        serverOutput: serverToClientPipe.Reader.AsStream()),
                    clientOptions: new McpClientOptions
                    {
                        ClientInfo = new() { Name = clientName, Version = "1.0.0" },
                        InitializationTimeout = ClientInitializationTimeout
                    },
                    cancellationToken: cancellationToken);
            }
            catch (Exception) when (DateTime.UtcNow < deadline && !cancellationToken.IsCancellationRequested)
            {
                await Task.Delay(ServerReadyRetryDelay, cancellationToken);
            }
        }
    }
}
