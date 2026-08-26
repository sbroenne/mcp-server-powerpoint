using System.ComponentModel;
using System.Diagnostics;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

internal static class OwnedProcessGuard
{
    internal static bool IsAlive(PowerPointProcessIdentity identity)
    {
        return TryOpenMatchingProcess(identity, out var process) && DisposeAndReturn(process);
    }

    internal static bool TryConfirmExited(PowerPointProcessIdentity identity)
    {
        return TryOpenMatchingProcess(identity, out var process) && !DisposeAndReturn(process);
    }

    internal static bool TryOpenMatchingProcess(
        PowerPointProcessIdentity identity,
        out Process? process)
    {
        process = null;
        try
        {
            var candidate = Process.GetProcessById(identity.ProcessId);
            if (candidate.HasExited
                || candidate.StartTime.ToUniversalTime().ToFileTimeUtc() != identity.StartedAtUtcFileTime)
            {
                candidate.Dispose();
                return true;
            }

            process = candidate;
            return true;
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (InvalidOperationException)
        {
            return true;
        }
        catch (Win32Exception)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
    }

    private static bool DisposeAndReturn(Process? process)
    {
        if (process == null)
        {
            return false;
        }

        process.Dispose();
        return true;
    }
}
