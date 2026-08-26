using System.Diagnostics;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class DaemonProcessTrackerTests
{
    [Fact]
    public void RegisterProcess_PersistsPipeScopedProcessIdentity()
    {
        var pipeName = UniquePipeName();
        using var current = Process.GetCurrentProcess();
        var startedAt = current.StartTime.ToUniversalTime().ToFileTimeUtc();

        try
        {
            var identity = DaemonProcessTracker.RegisterProcess(pipeName, current.Id, startedAt);
            var result = DaemonProcessTracker.ReadProcessSnapshot(pipeName);

            Assert.Equal(DaemonProcessTracker.TrackingRecordStatus.Available, result.Status);
            Assert.NotNull(result.Snapshot);
            Assert.Equal(identity, result.Snapshot.DaemonProcess);
            Assert.Empty(result.Snapshot.PowerPointProcesses);
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void TryOpenMatchingProcess_DoesNotMatchReusedPid()
    {
        using var current = Process.GetCurrentProcess();
        var staleIdentity = new DaemonProcessTracker.ProcessIdentity(
            current.Id,
            current.StartTime.ToUniversalTime().ToFileTimeUtc() - 1);

        Assert.True(DaemonProcessTracker.TryOpenMatchingProcess(staleIdentity, out var process));
        Assert.Null(process);
    }

    [Fact]
    public void ReadProcessSnapshot_RejectsMalformedRecord()
    {
        var pipeName = UniquePipeName();
        var path = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "{not-json");

            var result = DaemonProcessTracker.ReadProcessSnapshot(pipeName);

            Assert.Equal(DaemonProcessTracker.TrackingRecordStatus.Invalid, result.Status);
            Assert.Null(result.Snapshot);
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void RecordPowerPointProcesses_RejectsAnotherDaemonGeneration()
    {
        var pipeName = UniquePipeName();
        using var current = Process.GetCurrentProcess();
        var startedAt = current.StartTime.ToUniversalTime().ToFileTimeUtc();

        try
        {
            _ = DaemonProcessTracker.RegisterProcess(pipeName, current.Id, startedAt);
            var wrongGeneration = new DaemonProcessTracker.ProcessIdentity(current.Id, startedAt - 1);

            Assert.Throws<InvalidOperationException>(() =>
                DaemonProcessTracker.RecordPowerPointProcesses(pipeName, wrongGeneration, []));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void PipeIdentityHash_IsCaseInsensitive()
    {
        Assert.Equal(
            DaemonPipeIdentity.GetHash("PowerPointMcp-Test"),
            DaemonPipeIdentity.GetHash("powerpointmcp-test"));
    }

    [Fact]
    public void RecordPowerPointProcesses_AccumulatesOwnedIdentities()
    {
        var pipeName = UniquePipeName();
        using var current = Process.GetCurrentProcess();
        var startedAt = current.StartTime.ToUniversalTime().ToFileTimeUtc();

        try
        {
            var daemon = DaemonProcessTracker.RegisterProcess(pipeName, current.Id, startedAt);
            var first = new DaemonProcessTracker.ProcessIdentity(101, 1001);
            var second = new DaemonProcessTracker.ProcessIdentity(102, 1002);

            DaemonProcessTracker.RecordPowerPointProcesses(pipeName, daemon, [first]);
            DaemonProcessTracker.RecordPowerPointProcesses(pipeName, daemon, [first, second]);

            var result = DaemonProcessTracker.ReadProcessSnapshot(pipeName);
            Assert.Equal([first, second], result.Snapshot?.PowerPointProcesses);
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void RecordPowerPointProcess_RequiresValidDaemonRecord()
    {
        var tracker = new DaemonProcessTracker(UniquePipeName());
        var identity = new PowerPointProcessIdentity(101, 1001);

        Assert.Throws<InvalidOperationException>(() => tracker.RecordPowerPointProcess(identity));
    }

    private static string UniquePipeName() => $"PowerPointMcpTests-{Guid.NewGuid():N}";
}
