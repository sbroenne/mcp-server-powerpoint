using Sbroenne.PowerPointMcp.CLI.Commands;
using Sbroenne.PowerPointMcp.CLI.Generated;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.ComInterop.Session;
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
                session.SetDescription("Open, create, close, test, Save As/copy, or list presentation sessions held by the daemon; apply templates, manage the advisory Mark as Final flag, and read/write document properties.");
                session.AddCommand<SessionOpenCommand>("open").WithDescription("Open an existing presentation and return a session id.");
                session.AddCommand<SessionCreateCommand>("create").WithDescription("Create a new presentation and return a session id.");
                session.AddCommand<SessionCloseCommand>("close").WithDescription("Close a session, optionally saving first.");
                session.AddCommand<SessionListCommand>("list").WithDescription("List every session currently open in the daemon.");
                session.AddCommand<SessionTestCommand>("test").WithDescription("Validate that PowerPoint can open a presentation without retaining a session.");
                session.AddCommand<SessionSaveAsCommand>("save-as").WithDescription("Save the active presentation under a new path and move the session to it.");
                session.AddCommand<SessionSaveCopyAsCommand>("save-copy-as").WithDescription("Save a copy without changing the active presentation or session path.");
                session.AddCommand<SessionApplyTemplateCommand>("apply-template").WithDescription("Apply a template's masters/theme/layouts to the open presentation, preserving slide content.");
                session.AddCommand<SessionGetThemeNameCommand>("get-theme-name").WithDescription("Read the design/theme name currently applied to the open presentation.");
                session.AddCommand<SessionGetFinalCommand>("get-final").WithDescription("Read PowerPoint's advisory Mark as Final editing flag; it is not authentication, encryption, or access control.");
                session.AddCommand<SessionSetFinalCommand>("set-final").WithDescription("Set or clear PowerPoint's advisory Mark as Final editing flag; it is not authentication, encryption, or access control.");
                session.AddCommand<SessionSetDocumentPropertyCommand>("set-document-property").WithDescription("Set a built-in document metadata property (Title, Subject, Author, Keywords, Comments, Category, Manager, Company).");
                session.AddCommand<SessionGetDocumentPropertyCommand>("get-document-property").WithDescription("Read a built-in document metadata property.");
                session.AddCommand<SessionSetCustomPropertyCommand>("set-custom-property").WithDescription("Create or update a custom (user-defined) document property.");
                session.AddCommand<SessionGetCustomPropertyCommand>("get-custom-property").WithDescription("Read a custom (user-defined) document property.");
                session.AddCommand<SessionRemoveCustomPropertyCommand>("remove-custom-property").WithDescription("Remove a custom (user-defined) document property.");
                session.AddCommand<SessionSetTagCommand>("set-tag").WithDescription("Create or update a case-insensitive presentation string tag.");
                session.AddCommand<SessionGetTagCommand>("get-tag").WithDescription("Read a presentation string tag by case-insensitive name.");
                session.AddCommand<SessionListTagsCommand>("list-tags").WithDescription("List presentation string tags in native 1-based order.");
                session.AddCommand<SessionDeleteTagCommand>("delete-tag").WithDescription("Delete a presentation string tag by case-insensitive name.");
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
        using var daemonMutex = new Semaphore(
            initialCount: 1,
            maximumCount: 1,
            DaemonAutoStart.GetDaemonMutexName(pipeName));
        var daemonLockAcquired = false;
        try
        {
            daemonLockAcquired = daemonMutex.WaitOne(TimeSpan.Zero);

            if (!daemonLockAcquired)
            {
                Console.Error.WriteLine($"A daemon is already running for pipe '{pipeName}'.");
                return 1;
            }

            var tracker = new DaemonProcessTracker(pipeName);
            var trackingSnapshot = DaemonProcessTracker.ReadProcessSnapshot(pipeName);
            if (trackingSnapshot.Status == DaemonProcessTracker.TrackingRecordStatus.Invalid)
            {
                Console.Error.WriteLine(
                    $"The daemon tracking record is invalid and was preserved: {tracker.RecordPath}");
                return 1;
            }

            var staleRecord = tracker.ReadRecord();
            if (staleRecord != null
                && !OwnedProcessCleanup.TryCleanup(
                    staleRecord,
                    OperationDeadline.Start(TimeSpan.FromSeconds(10)),
                    out var cleanupError))
            {
                Console.Error.WriteLine(cleanupError);
                return 1;
            }

            if (staleRecord != null)
            {
                tracker.ClearIfMatches(staleRecord.Daemon);
            }

            var daemonIdentity = tracker.RegisterCurrentProcess();
            PresentationSessionRegistry.PowerPointProcessIdentityTracked += tracker.RecordPowerPointProcess;

            var service = new PowerPointMcpService();

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
                service.Dispose();
                PresentationSessionRegistry.PowerPointProcessIdentityTracked -= tracker.RecordPowerPointProcess;

                var finalRecord = tracker.ReadRecord();
                if (finalRecord != null && CanClearTrackingRecord(finalRecord))
                {
                    tracker.ClearIfMatches(daemonIdentity);
                }
            }

            return 0;
        }
        finally
        {
            if (daemonLockAcquired)
            {
                daemonMutex.Release();
            }
        }
    }

    private static bool CanClearTrackingRecord(DaemonTrackingRecord record)
    {
        foreach (var identity in record.PowerPointProcesses)
        {
            if (!DaemonProcessTracker.TryOpenMatchingProcess(identity, out var process))
            {
                return false;
            }

            using (process)
            {
                if (process != null)
                {
                    return false;
                }
            }
        }

        return true;
    }
}
