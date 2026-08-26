namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

internal static class OwnedProcessCleanup
{
    internal static bool TryTerminate(
        DaemonProcessIdentity identity,
        OperationDeadline deadline,
        out string? error)
    {
        error = null;
        if (!DaemonProcessTracker.TryOpenMatchingProcess(identity, out var process))
        {
            error = $"Could not validate ownership of process {identity.ProcessId}.";
            return false;
        }

        using (process)
        {
            if (process == null)
            {
                return true;
            }

            try
            {
                process.Kill(entireProcessTree: false);
                var remainingMilliseconds = (int)Math.Min(
                    int.MaxValue,
                    Math.Ceiling(deadline.Remaining.TotalMilliseconds));
                if (remainingMilliseconds <= 0 || !process.WaitForExit(remainingMilliseconds))
                {
                    error = $"Process {identity.ProcessId} did not exit before the cleanup deadline.";
                    return false;
                }

                return true;
            }
            catch (InvalidOperationException)
            {
                return true;
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                error = $"Could not terminate process {identity.ProcessId}: {ex.Message}";
                return false;
            }
        }
    }

    internal static bool TryCleanup(
        DaemonTrackingRecord record,
        OperationDeadline deadline,
        out string? error)
    {
        foreach (var identity in record.PowerPointProcesses)
        {
            if (!TryTerminate(identity, deadline, out error))
            {
                return false;
            }
        }

        return TryTerminate(record.Daemon, deadline, out error);
    }
}
