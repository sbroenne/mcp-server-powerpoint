using System.Diagnostics;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class DaemonProcessTrackerTests
{
    [Fact]
    public async Task TryGetLivePowerPointProcesses_LockedRecordReturnsWarning()
    {
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-locked-record-{Guid.NewGuid():N}.json");
        string mutexName = $@"Local\PowerPointMcp-Test-{Guid.NewGuid():N}";
        var tracker = new DaemonProcessTracker(path, mutexName);
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var release = new ManualResetEventSlim();
        var holder = Task.Run(() =>
        {
            using var mutex = new Mutex(false, mutexName);
            Assert.True(mutex.WaitOne(TimeSpan.FromSeconds(5)));
            try
            {
                started.SetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        });
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(tracker.TryGetLivePowerPointProcesses(out var processes, out var error));
            Assert.Empty(processes);
            Assert.Contains("tracking record", error, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            release.Set();
            await holder.WaitAsync(TimeSpan.FromSeconds(5));
            File.Delete(path);
        }
    }

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

    [Fact]
    public void Recovery_DoesNotTerminateRecordedPowerPointProcess()
    {
        using var child = Process.Start(new ProcessStartInfo("pwsh", "-NoProfile -Command \"Start-Sleep 60\"")
        {
            UseShellExecute = false,
            CreateNoWindow = true
        }) ?? throw new InvalidOperationException("Could not start cleanup test process.");
        using var current = Process.GetCurrentProcess();
        var record = new DaemonTrackingRecord
        {
            Daemon = new DaemonProcessIdentity(current.Id, current.StartTime.ToUniversalTime().ToFileTimeUtc() - 1),
            PowerPointProcesses =
            [
                new DaemonProcessIdentity(child.Id, child.StartTime.ToUniversalTime().ToFileTimeUtc())
            ]
        };
        try
        {
            Assert.True(OwnedProcessCleanup.TryCleanup(record, OperationDeadline.Start(TimeSpan.FromSeconds(3)), out var error), error);
            Assert.False(child.HasExited, "Recovering a CLI service must not end PowerPoint or lose unsaved work.");
        }
        finally
        {
            if (!child.HasExited)
            {
                child.Kill();
                child.WaitForExit(5000);
            }
        }
    }

    [Fact]
    public void RegisterCurrentProcess_PreservesEarlierPowerPointRecords()
    {
        string pipeName = UniquePipeName();
        using var current = Process.GetCurrentProcess();
        var tracker = new DaemonProcessTracker(pipeName);
        try
        {
            var daemon = DaemonProcessTracker.RegisterProcess(pipeName, current.Id, current.StartTime.ToUniversalTime().ToFileTimeUtc());
            var earlier = new DaemonProcessTracker.ProcessIdentity(101, 1001);
            DaemonProcessTracker.RecordPowerPointProcesses(pipeName, daemon, [earlier]);

            tracker.RegisterCurrentProcess();

            Assert.Contains(earlier, DaemonProcessTracker.ReadProcessSnapshot(pipeName).Snapshot!.PowerPointProcesses);
        }
        finally
        {
            tracker.Clear();
        }
    }

    private static string UniquePipeName() => $"PowerPointMcpTests-{Guid.NewGuid():N}";
}
