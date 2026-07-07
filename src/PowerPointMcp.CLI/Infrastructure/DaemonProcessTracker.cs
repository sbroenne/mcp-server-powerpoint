using System.Diagnostics;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Tracks the daemon process id in a small per-pipe-name file so <c>service stop --force</c> can
/// kill a stuck daemon that isn't responding to a graceful RPC shutdown. Simplified relative to
/// mcp-server-excel's DaemonProcessTracker (no WinForms tray integration — see squad decision
/// 2026-07-07 for the scope note on the tray icon being dropped for this port).
/// </summary>
internal static class DaemonProcessTracker
{
    private static string GetTrackerFilePath(string pipeName) =>
        Path.Combine(Path.GetTempPath(), $"powerpointmcp-daemon-{pipeName}.pid");

    /// <summary>Records the daemon's process id for later force-kill lookup.</summary>
    public static void Track(string pipeName, int processId)
    {
        try
        {
            File.WriteAllText(GetTrackerFilePath(pipeName), processId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        catch (IOException)
        {
            // Best-effort — force-stop simply won't have a tracked PID to fall back on.
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>Removes the tracked process id (daemon exited on its own).</summary>
    public static void Clear(string pipeName)
    {
        try
        {
            File.Delete(GetTrackerFilePath(pipeName));
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>
    /// Attempts to kill a previously-tracked daemon process directly, for use when a graceful
    /// RPC shutdown does not work (daemon stuck/unresponsive).
    /// </summary>
    /// <returns><see langword="true"/> if a tracked process was found and killed.</returns>
    public static bool TryForceStopTrackedDaemon(string pipeName)
    {
        var path = GetTrackerFilePath(pipeName);
        try
        {
            if (!File.Exists(path) || !int.TryParse(File.ReadAllText(path), out var pid))
            {
                return false;
            }

            using var process = Process.GetProcessById(pid);
            process.Kill(entireProcessTree: true);
            process.WaitForExit(5000);
            return true;
        }
        catch (ArgumentException)
        {
            // No process with that id — already gone.
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        finally
        {
            Clear(pipeName);
        }
    }
}
