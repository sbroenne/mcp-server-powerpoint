using System.Diagnostics;
using System.Text.Json;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

internal sealed class DaemonProcessTracker
{
    internal sealed record ProcessIdentity(int ProcessId, long StartedAtUtcFileTime);

    internal sealed record ProcessSnapshot(
        ProcessIdentity DaemonProcess,
        IReadOnlyList<ProcessIdentity> PowerPointProcesses);

    internal enum TrackingRecordStatus
    {
        Missing,
        Available,
        Invalid,
    }

    internal sealed record ProcessSnapshotResult(
        TrackingRecordStatus Status,
        ProcessSnapshot? Snapshot);

    private readonly string _recordPath;
    private readonly string _mutexName;

    public DaemonProcessTracker(string? pipeName = null)
    {
        var pipeKey = DaemonPipeIdentity.GetStableKey(pipeName);
        var directory = Path.Combine(Path.GetTempPath(), "PowerPointMcp", "cli-daemon");
        _recordPath = Path.Combine(directory, $"{pipeKey}.json");
        _mutexName = $@"Global\PowerPointMcp-DaemonRecord-{pipeKey}";
    }

    internal DaemonProcessTracker(string recordPath, string mutexName)
    {
        _recordPath = recordPath;
        _mutexName = mutexName;
    }

    internal string RecordPath => _recordPath;

    internal static string GetTrackingFilePath(string pipeName) =>
        new DaemonProcessTracker(pipeName).RecordPath;

    internal static ProcessIdentity RegisterProcess(
        string pipeName,
        int processId,
        long startedAtUtcFileTime)
    {
        var tracker = new DaemonProcessTracker(pipeName);
        var identity = new DaemonProcessIdentity(processId, startedAtUtcFileTime);
        tracker.WithLock(() => tracker.WriteRecord(new DaemonTrackingRecord
        {
            Daemon = identity,
            PowerPointProcesses = [],
        }));
        return new ProcessIdentity(processId, startedAtUtcFileTime);
    }

    internal static ProcessSnapshotResult ReadProcessSnapshot(string pipeName)
    {
        var tracker = new DaemonProcessTracker(pipeName);
        if (!File.Exists(tracker.RecordPath))
        {
            return new ProcessSnapshotResult(TrackingRecordStatus.Missing, null);
        }

        var record = tracker.ReadRecord();
        if (record == null)
        {
            return new ProcessSnapshotResult(TrackingRecordStatus.Invalid, null);
        }

        return new ProcessSnapshotResult(
            TrackingRecordStatus.Available,
            new ProcessSnapshot(
                new ProcessIdentity(record.Daemon.ProcessId, record.Daemon.StartedAtUtcFileTime),
                record.PowerPointProcesses
                    .Select(identity => new ProcessIdentity(identity.ProcessId, identity.StartedAtUtcFileTime))
                    .ToArray()));
    }

    internal static void RecordPowerPointProcesses(
        string pipeName,
        ProcessIdentity daemonGeneration,
        IReadOnlyList<ProcessIdentity> powerPointProcesses)
    {
        var tracker = new DaemonProcessTracker(pipeName);
        tracker.WithLock(() =>
        {
            var record = tracker.ReadRecordCore()
                ?? throw new InvalidOperationException("The daemon tracking record is missing or invalid.");
            if (record.Daemon.ProcessId != daemonGeneration.ProcessId
                || record.Daemon.StartedAtUtcFileTime != daemonGeneration.StartedAtUtcFileTime)
            {
                throw new InvalidOperationException(
                    "The PowerPoint process identities belong to another daemon generation.");
            }

            foreach (var process in powerPointProcesses)
            {
                var identity = new DaemonProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime);
                if (!record.PowerPointProcesses.Contains(identity))
                {
                    record.PowerPointProcesses.Add(identity);
                }
            }

            tracker.WriteRecord(record);
        });
    }

    internal static void Clear(string pipeName) =>
        new DaemonProcessTracker(pipeName).Clear();

    public DaemonProcessIdentity RegisterCurrentProcess()
    {
        using var process = Process.GetCurrentProcess();
        var identity = new DaemonProcessIdentity(
            process.Id,
            process.StartTime.ToUniversalTime().ToFileTimeUtc());

        WithLock(() => WriteRecord(new DaemonTrackingRecord
        {
            Daemon = identity,
            PowerPointProcesses = [],
        }));
        return identity;
    }

    public void RecordPowerPointProcess(PowerPointProcessIdentity identity)
    {
        WithLock(() =>
        {
            var record = ReadRecordCore();
            if (record == null)
            {
                throw new InvalidOperationException(
                    "Cannot record the PowerPoint process because the daemon tracking record is missing or invalid.");
            }

            var process = new DaemonProcessIdentity(identity.ProcessId, identity.StartedAtUtcFileTime);
            if (!record.PowerPointProcesses.Contains(process))
            {
                record.PowerPointProcesses.Add(process);
                WriteRecord(record);
            }
        });
    }

    public DaemonTrackingRecord? ReadRecord() =>
        WithLock(ReadRecordCore);

    public void Clear() =>
        WithLock(DeleteRecordCore);

    public void ClearIfMatches(DaemonProcessIdentity identity)
    {
        WithLock(() =>
        {
            if (ReadRecordCore()?.Daemon == identity)
            {
                DeleteRecordCore();
            }
        });
    }

    public static bool TryOpenMatchingProcess(
        DaemonProcessIdentity identity,
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
        catch (System.ComponentModel.Win32Exception)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
    }

    internal static bool TryOpenMatchingProcess(
        ProcessIdentity identity,
        out Process? process) =>
        TryOpenMatchingProcess(
            new DaemonProcessIdentity(identity.ProcessId, identity.StartedAtUtcFileTime),
            out process);

    private DaemonTrackingRecord? ReadRecordCore()
    {
        if (!File.Exists(_recordPath))
        {
            return null;
        }

        try
        {
            var record = JsonSerializer.Deserialize<DaemonTrackingRecord>(
                File.ReadAllText(_recordPath),
                DaemonTrackingJson.Options);
            return record is { Daemon.ProcessId: > 0, Daemon.StartedAtUtcFileTime: > 0 }
                ? record
                : null;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private void WriteRecord(DaemonTrackingRecord record)
    {
        var directory = Path.GetDirectoryName(_recordPath)
            ?? throw new InvalidOperationException("Daemon tracking path has no directory.");
        Directory.CreateDirectory(directory);

        var temporaryPath = _recordPath + $".{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(
                temporaryPath,
                JsonSerializer.Serialize(record, DaemonTrackingJson.Options));
            File.Move(temporaryPath, _recordPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }

    private void DeleteRecordCore()
    {
        if (File.Exists(_recordPath))
        {
            File.Delete(_recordPath);
        }
    }

    private void WithLock(Action action)
    {
        WithLock(() =>
        {
            action();
            return true;
        });
    }

    private T WithLock<T>(Func<T> action)
    {
        using var mutex = new Mutex(false, _mutexName);
        var lockTaken = false;
        try
        {
            try
            {
                lockTaken = mutex.WaitOne(TimeSpan.FromSeconds(5));
            }
            catch (AbandonedMutexException)
            {
                lockTaken = true;
            }

            if (!lockTaken)
            {
                throw new TimeoutException("Timed out while accessing the daemon tracking record.");
            }

            return action();
        }
        finally
        {
            if (lockTaken)
            {
                mutex.ReleaseMutex();
            }
        }
    }
}

internal sealed record DaemonProcessIdentity(
    int ProcessId,
    long StartedAtUtcFileTime);

internal sealed class DaemonTrackingRecord
{
    public required DaemonProcessIdentity Daemon { get; init; }

    public List<DaemonProcessIdentity> PowerPointProcesses { get; init; } = [];
}
