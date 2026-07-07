using Sbroenne.PowerPointMcp.CLI.Commands;
using Sbroenne.PowerPointMcp.CLI.Generated;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI;

/// <summary>Entry point for the <c>pptcli</c> command-line tool.</summary>
public static class Program
{
    /// <summary>
    /// Runs the CLI. A special <c>service run</c> invocation launches the daemon in-process
    /// (blocking); every other invocation goes through Spectre.Console.Cli's command app, which
    /// dispatches to the daemon over a named pipe, auto-starting it on first use.
    /// </summary>
    public static async Task<int> Main(string[] args)
    {
        if (args.Length >= 2 && args[0] == "service" && args[1] == "run")
        {
            return await RunDaemonAsync(args);
        }

        var app = new CommandApp();
        app.Configure(config =>
        {
            config.SetApplicationName("pptcli");

            config.AddBranch("session", session =>
            {
                session.SetDescription("Open, create, save, close, or list presentation sessions held by the daemon.");
                session.AddCommand<SessionOpenCommand>("open").WithDescription("Open an existing presentation and return a session id.");
                session.AddCommand<SessionCreateCommand>("create").WithDescription("Create a new presentation and return a session id.");
                session.AddCommand<SessionCloseCommand>("close").WithDescription("Close a session, optionally saving first.");
                session.AddCommand<SessionSaveCommand>("save").WithDescription("Save the presentation open in a session.");
                session.AddCommand<SessionListCommand>("list").WithDescription("List every session currently open in the daemon.");
            });

            config.AddBranch("service", service =>
            {
                service.SetDescription("Start, stop, or check the status of the pptcli background daemon.");
                service.AddCommand<ServiceStartCommand>("start").WithDescription("Start the daemon if it isn't already running.");
                service.AddCommand<ServiceStopCommand>("stop").WithDescription("Stop the running daemon.");
                service.AddCommand<ServiceStatusCommand>("status").WithDescription("Report whether the daemon is running.");
            });

            CliCommandRegistration.RegisterCommands(config);
        });

        return await app.RunAsync(args);
    }

    /// <summary>
    /// Runs the daemon in the current process until it shuts down (idle timeout, explicit
    /// <c>service stop</c>, or Ctrl+C). This is the process launched by
    /// <see cref="DaemonAutoStart"/> when no daemon is currently listening on the pipe.
    /// </summary>
    private static async Task<int> RunDaemonAsync(string[] args)
    {
        string? pipeName = null;
        var idleTimeout = TimeSpan.FromMinutes(10);

        for (var i = 2; i < args.Length; i++)
        {
            if (args[i] == "--pipe-name" && i + 1 < args.Length)
            {
                pipeName = args[++i];
            }
            else if (args[i] == "--idle-timeout-minutes" && i + 1 < args.Length && double.TryParse(args[i + 1], out var minutes))
            {
                idleTimeout = TimeSpan.FromMinutes(minutes);
                i++;
            }
        }

        pipeName ??= DaemonAutoStart.GetPipeName();

        using var service = new PowerPointMcpService();

        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            service.RequestShutdown();
        };

        try
        {
            await service.RunAsync(pipeName, idleTimeout);
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }

        return 0;
    }
}
