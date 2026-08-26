using System.Diagnostics;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.ComInterop.Tests;

public sealed class OwnedProcessGuardTests
{
    [Fact]
    public void IsAlive_RejectsSamePidWithDifferentStartTime()
    {
        using var current = Process.GetCurrentProcess();
        var identity = new PowerPointProcessIdentity(
            current.Id,
            current.StartTime.ToUniversalTime().ToFileTimeUtc() - 1);

        Assert.False(OwnedProcessGuard.IsAlive(identity));
    }

    [Fact]
    public void IsAlive_MatchesCurrentProcessIdentity()
    {
        using var current = Process.GetCurrentProcess();
        var identity = new PowerPointProcessIdentity(
            current.Id,
            current.StartTime.ToUniversalTime().ToFileTimeUtc());

        Assert.True(OwnedProcessGuard.IsAlive(identity));
    }
}
