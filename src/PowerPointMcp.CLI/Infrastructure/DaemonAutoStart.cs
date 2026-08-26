using System.Diagnostics;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Ensures the CLI daemon is running before sending commands, auto-starting it if necessary.
/// Ported from mcp-server-excel's ExcelMcp.CLI.Infrastructure.DaemonAutoStart.
/// </summary>
internal static class DaemonAutoStart
{
    internal static readonly TimeSpan InitialPingTimeout = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan BusyDaemonConnectTimeout = TimeSpan.FromSeconds(3);
    internal static readonly TimeSpan BusyDaemonRetryInterval = TimeSpan.FromMilliseconds(500);
    internal static readonly TimeSpan BusyDaemonWaitTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan StartupReadyConnectTimeout = TimeSpan.FromSeconds(1);
    internal static readonly TimeSpan StartupReadyRetryInterval = TimeSpan.FromMilliseconds(250);
    internal static readonly TimeSpan StartupReadyTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan StartupLockTimeout = StartupReadyTimeout + TimeSpan.FromSeconds(1);

    /// <summary>Gets the pipe name for the CLI daemon (supports env var override for testing).</summary>
    public static string GetPipeName() =>
        Environment.GetEnvironmentVariable("POWERPOINTMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();

    /// <summary>
    /// Ensures the CLI daemon is running and returns a connected <see cref="ServiceClient"/>.
    /// If the daemon is not running, starts it and waits for it to be ready.
    /// </summary>
    public static async Task<ServiceClient> EnsureAndConnectAsync(CancellationToken cancellationToken = default)
    {
        var pipeName = GetPipeName();

        // Fast path: daemon already running and responsive.
        if (await PingAsync(pipeName, InitialPingTimeout, cancellationToken))
        {
            return new ServiceClient(pipeName);
        }

        // Ping failed — check the OS mutex to distinguish "daemon busy" from "daemon not
        // running". The daemon holds this mutex for its entire lifetime.
        if (IsDaemonMutexHeld(pipeName))
        {
            var deadline = OperationDeadline.Start(BusyDaemonWaitTimeout);
            while (!deadline.IsExpired)
            {
                await Task.Delay(BusyDaemonRetryInterval, cancellationToken);

                if (!IsDaemonMutexHeld(pipeName))
                    break;

                var remaining = deadline.Remaining;
                if (remaining <= TimeSpan.Zero)
                    break;

                if (await PingAsync(pipeName, Min(remaining, BusyDaemonConnectTimeout), cancellationToken))
                    return new ServiceClient(pipeName);
            }

            if (IsDaemonMutexHeld(pipeName))
            {
                throw new TimeoutException(
                    $"Daemon is running but not responding after {FormatDuration(BusyDaemonWaitTimeout)}. " +
                    "Stop it with 'pptcli service stop' or terminate the stuck pptcli process, then retry.");
            }

            // Daemon exited while we waited — start a replacement.
        }

        var startupDeadline = OperationDeadline.Start(StartupLockTimeout);
        if (!await TryStartDaemonWithStartupLockAsync(pipeName, startupDeadline, cancellationToken))
        {
            if (await WaitForResponsiveDaemonAsync(pipeName, startupDeadline.Remaining, cancellationToken))
                return new ServiceClient(pipeName);

            throw new TimeoutException(
                $"Daemon startup is already in progress but did not become ready within {FormatDuration(StartupReadyTimeout)}.");
        }

        if (!startupDeadline.IsExpired
            && await PingAsync(pipeName, startupDeadline.Cap(StartupReadyConnectTimeout), cancellationToken))
        {
            return new ServiceClient(pipeName);
        }

        throw new TimeoutException($"Daemon started but not responding within {FormatDuration(StartupReadyTimeout)}.");
    }

    /// <summary>
    /// Checks whether a daemon process currently holds the daemon mutex for the given pipe name.
    /// </summary>
    internal static bool IsDaemonMutexHeld(string pipeName)
    {
        Semaphore? semaphore = null;
        try
        {
            semaphore = Semaphore.OpenExisting(GetDaemonMutexName(pipeName));
            if (semaphore.WaitOne(TimeSpan.Zero))
            {
                semaphore.Release();
                return false;
            }

            return true;
        }
        catch (WaitHandleCannotBeOpenedException)
        {
            return false; // No process has this mutex — daemon is not running.
        }
        catch (UnauthorizedAccessException)
        {
            return true; // Fail closed: inability to inspect the lock must not start a second daemon.
        }
        finally
        {
            semaphore?.Dispose();
        }
    }

    internal static bool IsDaemonStartupMutexHeld(string pipeName) =>
        IsMutexHeld(GetDaemonStartupLockName(pipeName));

    private static bool IsMutexHeld(string mutexName)
    {
        Mutex? mutex = null;
        try
        {
            mutex = Mutex.OpenExisting(mutexName);
            if (!mutex.WaitOne(TimeSpan.Zero))
            {
                return true;
            }

            mutex.ReleaseMutex();
            return false;
        }
        catch (AbandonedMutexException)
        {
            try
            {
                mutex?.ReleaseMutex();
            }
            catch (ApplicationException)
            {
            }

            return false;
        }
        catch (WaitHandleCannotBeOpenedException)
        {
            return false;
        }
        finally
        {
            mutex?.Dispose();
        }
    }

    /// <summary>
    /// Gets the OS mutex name for the CLI daemon identified by its pipe name. Used by both the
    /// daemon (to acquire) and the client (to detect a running daemon).
    /// </summary>
    internal static string GetDaemonMutexName(string pipeName) =>
        $"PowerPointMcpCli_{pipeName}";

    internal static string GetDaemonStartupLockName(string pipeName) =>
        $"{GetDaemonMutexName(pipeName)}_startup";

    private static Task<bool> TryStartDaemonWithStartupLockAsync(
        string pipeName,
        OperationDeadline deadline,
        CancellationToken cancellationToken)
    {
        return Task.Run(() =>
        {
            using var startupMutex = new Mutex(initiallyOwned: false, GetDaemonStartupLockName(pipeName), out _);
            var startupLockAcquired = false;
            try
            {
                try
                {
                    startupLockAcquired = startupMutex.WaitOne(deadline.Cap(StartupLockTimeout));
                }
                catch (AbandonedMutexException)
                {
                    startupLockAcquired = true;
                }

                if (!startupLockAcquired)
                    return false;

                if (deadline.IsExpired)
                    return false;

                // Another CLI process may have started the daemon while this process waited.
                if (PingAsync(pipeName, deadline.Cap(InitialPingTimeout), cancellationToken).GetAwaiter().GetResult())
                    return true;

                StartDaemonAsync(pipeName, deadline, cancellationToken).GetAwaiter().GetResult();
                return true;
            }
            finally
            {
                if (startupLockAcquired)
                    startupMutex.ReleaseMutex();
            }
        }, cancellationToken);
    }

    private static async Task<bool> WaitForResponsiveDaemonAsync(string pipeName, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = OperationDeadline.Start(timeout);
        while (!deadline.IsExpired)
        {
            await Task.Delay(StartupReadyRetryInterval, cancellationToken);

            var remaining = deadline.Remaining;
            if (remaining <= TimeSpan.Zero)
                break;

            if (await PingAsync(pipeName, Min(remaining, StartupReadyConnectTimeout), cancellationToken))
                return true;
        }

        return false;
    }

    private static async Task StartDaemonAsync(
        string pipeName,
        OperationDeadline deadline,
        CancellationToken cancellationToken)
    {
        var exePath = ResolveDaemonExecutablePath();

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"service run --pipe-name \"{pipeName}\"",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(exePath) ?? Environment.CurrentDirectory
        };

        try
        {
            using var daemonProcess = Process.Start(startInfo)
                ?? throw new InvalidOperationException($"Failed to start daemon process '{exePath}'.");

            while (!deadline.IsExpired)
            {
                await Task.Delay(StartupReadyRetryInterval, cancellationToken);
                if (daemonProcess.HasExited)
                {
                    if (daemonProcess.ExitCode == 0)
                    {
                        if (await WaitForResponsiveDaemonAsync(pipeName, deadline.Remaining, cancellationToken))
                        {
                            GC.KeepAlive(daemonProcess);
                            return;
                        }

                        throw new InvalidOperationException(
                            "Daemon process exited cleanly before becoming ready, but no responsive daemon was found. " +
                            "This usually means a stale startup race or a daemon that shut down immediately. " +
                            "Run 'pptcli service stop' and retry.");
                    }

                    throw new InvalidOperationException(
                        $"Daemon process exited before becoming ready (exit code {daemonProcess.ExitCode}).");
                }

                var remaining = deadline.Remaining;
                if (remaining <= TimeSpan.Zero)
                    break;

                if (await PingAsync(pipeName, deadline.Cap(StartupReadyConnectTimeout), cancellationToken))
                {
                    GC.KeepAlive(daemonProcess);
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to start daemon: {ex.Message}", ex);
        }

        throw new TimeoutException($"Daemon started but not responding within {FormatDuration(StartupReadyTimeout)}.");
    }

    private static string ResolveDaemonExecutablePath()
    {
        var baseDirectoryCandidate = Path.Combine(AppContext.BaseDirectory, "powerpointcli.exe");
        if (File.Exists(baseDirectoryCandidate))
        {
            return baseDirectoryCandidate;
        }

        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath) && File.Exists(processPath))
        {
            return processPath;
        }

        throw new InvalidOperationException("Cannot determine executable path to start daemon.");
    }

    private static string FormatDuration(TimeSpan duration)
    {
        return duration.TotalSeconds >= 1
            ? $"{duration.TotalSeconds:0.#} seconds"
            : $"{duration.TotalMilliseconds:0} ms";
    }

    private static TimeSpan Min(TimeSpan left, TimeSpan right) => left <= right ? left : right;

    private static async Task<bool> PingAsync(string pipeName, TimeSpan connectTimeout, CancellationToken cancellationToken)
    {
        var requestTimeout = connectTimeout + TimeSpan.FromSeconds(1);
        using var client = new ServiceClient(pipeName, connectTimeout: connectTimeout, requestTimeout: requestTimeout);
        return await client.PingAsync(cancellationToken);
    }
}
