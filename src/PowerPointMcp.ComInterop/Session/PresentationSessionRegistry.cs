using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Lightweight, immutable view of a single registered presentation session.
/// Returned by <see cref="PresentationSessionRegistry.List"/> for tooling that needs to
/// enumerate open sessions without touching the underlying COM batch directly.
/// </summary>
/// <param name="SessionId">The generated identifier used to look the session up.</param>
/// <param name="PresentationPath">Full path to the presentation the batch is bound to.</param>
/// <param name="IsPowerPointProcessAlive">Whether the backing POWERPNT.exe process is still alive.</param>
public sealed record PresentationSessionInfo(
    string SessionId,
    string PresentationPath,
    bool IsPowerPointProcessAlive)
{
    /// <summary>Last operation state, read without calling PowerPoint.</summary>
    public PresentationOperationStatus? Operation { get; init; }
}

/// <summary>
/// In-process registry that maps a generated session identifier to a long-lived
/// <see cref="IPresentationBatch"/>. Registered as a DI singleton for the lifetime of the
/// MCP host so that one PowerPoint session can span many tool invocations.
/// </summary>
/// <remarks>
/// Used by the MCP Server as an in-process singleton (one long-lived session per tool-call
/// sequence, no RPC needed since the MCP host and the STA thread live in the same process).
/// Also reused by <c>PowerPointMcp.Service</c> (the CLI's out-of-process daemon) as the daemon's
/// session table, mirroring mcp-server-excel's <c>SessionManager</c> — there each session maps to a live
/// <c>IPresentationBatch</c> held open across separate CLI process invocations, avoiding a
/// PowerPoint relaunch on every command. Disposing a batch closes PowerPoint for that
/// presentation, so hosts must call <see cref="DisposeAll()"/> on shutdown to release automation
/// while leaving unsaved presentations open.
///
/// Close removes a session immediately and tracks its background cleanup. Host shutdown waits
/// only for bounded cleanup, never for PowerPoint's process to exit, and rejects later startup.
/// </remarks>
public sealed class PresentationSessionRegistry : IDisposable
{
    /// <summary>
    /// Overall bound for <see cref="DisposeAll()"/> to wait on in-flight background disposals.
    /// Sized above <see cref="ComInteropConstants.StaThreadJoinTimeout"/> (the worst-case time a
    /// single batch's Dispose() can take) plus a buffer, since concurrently-closing sessions
    /// dispose in parallel rather than sequentially.
    /// </summary>
    private static readonly TimeSpan DisposeAllTimeout = ComInteropConstants.StaThreadJoinTimeout + TimeSpan.FromSeconds(30);

    private static readonly ConcurrentDictionary<PowerPointProcessIdentity, int> TrackedPowerPointProcesses = new();
    private static readonly object ProcessOwnershipLock = new();
    private static int _processExitRegistered;

    /// <summary>Raised when this host captures a new owned PowerPoint process identity.</summary>
    internal static event Action<PowerPointProcessIdentity>? PowerPointProcessIdentityTracked;

    internal static IReadOnlyList<PowerPointProcessIdentity> GetTrackedPowerPointProcesses() =>
        TrackedPowerPointProcesses.Keys.ToArray();

    /// <summary>Registers an owned PowerPoint identity for crash and daemon cleanup.</summary>
    internal static void TrackPowerPointProcess(PowerPointProcessIdentity identity)
    {
        lock (ProcessOwnershipLock)
            TrackedPowerPointProcesses.AddOrUpdate(identity, 1, (_, owners) => owners + 1);
        PowerPointProcessIdentityTracked?.Invoke(identity);

        if (Interlocked.CompareExchange(ref _processExitRegistered, 1, 0) == 0)
        {
            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
        }
    }

    internal static void ReleasePowerPointProcessOwnership(PowerPointProcessIdentity identity)
    {
        lock (ProcessOwnershipLock)
        {
            if (!TrackedPowerPointProcesses.TryGetValue(identity, out var owners))
                return;
            if (owners > 1)
                TrackedPowerPointProcesses[identity] = owners - 1;
            else
                TrackedPowerPointProcesses.TryRemove(identity, out _);
        }
    }

    private static void OnProcessExit(object? sender, EventArgs e)
    {
        foreach (var identity in TrackedPowerPointProcesses.Keys)
        {
            PresentationDiagnostics.WriteStdErr(
                $"[DIAG-PROCESSEXIT-RETAINED] PowerPoint process {identity.ProcessId} was not terminated; unsaved work may remain open.");
        }
    }

    private readonly ConcurrentDictionary<string, IPresentationBatch> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<Task, byte> _pendingDisposals = new();
    private readonly object _lifecycleLock = new();
    private readonly ILogger<PresentationSessionRegistry>? _logger;
    private bool _disposed;

    /// <summary>Creates a registry with no logging.</summary>
    public PresentationSessionRegistry()
        : this(null)
    {
    }

    /// <summary>Creates a registry that logs lifecycle diagnostics to the provided logger.</summary>
    public PresentationSessionRegistry(ILogger<PresentationSessionRegistry>? logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Opens an existing presentation file and registers the resulting batch.
    /// </summary>
    /// <param name="filePath">Full path to an existing .pptx/.pptm/.ppt file.</param>
    /// <returns>The generated session identifier.</returns>
    public string Open(string filePath)
    {
        lock (_lifecycleLock)
            ObjectDisposedException.ThrowIf(_disposed, this);
        var batch = PresentationSession.BeginBatch(filePath);
        return Register(batch);
    }

    /// <summary>
    /// Creates a new presentation file, leaves it open, and registers the resulting batch.
    /// </summary>
    /// <param name="filePath">Full path to the new .pptx/.pptm file to create.</param>
    /// <returns>The generated session identifier.</returns>
    public string Create(string filePath)
    {
        lock (_lifecycleLock)
            ObjectDisposedException.ThrowIf(_disposed, this);
        var batch = PresentationSession.CreateNew(filePath);
        return Register(batch);
    }

    /// <summary>
    /// Attempts to resolve a session identifier to its live batch.
    /// </summary>
    /// <param name="sessionId">The identifier returned by <see cref="Open"/> or <see cref="Create"/>.</param>
    /// <param name="batch">The resolved batch when the method returns <see langword="true"/>.</param>
    /// <returns><see langword="true"/> if a session with the given id exists.</returns>
    public bool TryGet(string sessionId, out IPresentationBatch batch)
    {
        if (!string.IsNullOrEmpty(sessionId) && _sessions.TryGetValue(sessionId, out var found))
        {
            batch = found;
            return true;
        }

        batch = null!;
        return false;
    }

    /// <summary>
    /// Closes a session: atomically removes it from the registry and starts disposing its batch
    /// on a background task. Returns immediately, without waiting for PowerPoint to exit. The
    /// background dispose is tracked so <see cref="DisposeAll()"/> can await it on host shutdown.
    /// No-op if the identifier is unknown or another request is still running. Reserving the
    /// batch for close prevents work from starting through an earlier lookup.
    /// </summary>
    /// <param name="sessionId">The session identifier to close.</param>
    /// <param name="pendingOperation">
    /// A cancelled request's still-running operation. Remove its session immediately, but wait for
    /// the operation to finish before disposing the batch. The caller observes operation failures.
    /// </param>
    /// <returns><see langword="true"/> if a session was found and its close was started; otherwise, unknown or busy.</returns>
    public bool Close(string sessionId, Task? pendingOperation = null, bool preserveUnsavedChanges = false)
    {
        lock (_lifecycleLock)
        {
            if (string.IsNullOrEmpty(sessionId) || !_sessions.TryGetValue(sessionId, out var batch))
                return false;

            if (!preserveUnsavedChanges && pendingOperation is null && !batch.TryBeginClose())
                return false;
            _sessions.TryRemove(sessionId, out _);
            StartBackgroundDispose(sessionId, batch, pendingOperation, preserveUnsavedChanges);
            return true;
        }
    }

    /// <summary>
    /// Returns metadata for every currently open session.
    /// </summary>
    public IReadOnlyList<PresentationSessionInfo> List()
    {
        var result = new List<PresentationSessionInfo>(_sessions.Count);
        foreach (var pair in _sessions)
        {
            result.Add(new PresentationSessionInfo(
                pair.Key,
                pair.Value.PresentationPath,
                pair.Value.IsPowerPointProcessAlive())
            {
                Operation = pair.Value.GetOperationStatus()
            });
        }

        return result;
    }

    /// <summary>
    /// Removes every registered session and starts disposing each batch on a background task,
    /// then blocks (up to <see cref="DisposeAllTimeout"/>) until every in-flight disposal —
    /// including any still pending from earlier <see cref="Close"/> calls — has finished. Call
    /// this from host shutdown so the process does not exit while a POWERPNT.exe is still being
    /// cleaned up. Idempotent and safe to call repeatedly (e.g. from both the hosted-service
    /// StopAsync and Main's finally backstop). Never throws — failures are logged.
    /// </summary>
    public void DisposeAll() => DisposeAll(DisposeAllTimeout);

    /// <summary>
    /// Overload used by tests to bound the wait with a shorter timeout than production. See
    /// <see cref="DisposeAll()"/> for behavior.
    /// </summary>
    internal void DisposeAll(TimeSpan overallTimeout)
    {
        lock (_lifecycleLock)
        {
            _disposed = true;
            foreach (var sessionId in _sessions.Keys.ToArray())
            {
                if (_sessions.TryRemove(sessionId, out var batch))
                    StartBackgroundDispose(sessionId, batch, preserveUnsavedChanges: true);
            }
        }

        // Snapshot AFTER starting our own disposals so this call also waits for anything a
        // concurrent/earlier Close() kicked off and hasn't finished yet.
        var pending = _pendingDisposals.Keys.ToArray();
        if (pending.Length == 0)
        {
            return;
        }

        try
        {
            if (!Task.WhenAll(pending).Wait(overallTimeout))
            {
                _logger?.LogWarning(
                    "Stopped waiting after {TimeoutSeconds:N0}s for {Count} PowerPoint session(s) to release automation. PowerPoint was not terminated; unsaved work may remain open.",
                    overallTimeout.TotalSeconds,
                    pending.Length);
            }
        }
#pragma warning disable CA1031 // Shutdown wait must never throw out of DisposeAll.
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Unexpected error while waiting for background PowerPoint session disposal during shutdown.");
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        lock (_lifecycleLock)
        {
            if (_disposed)
                return;
            _disposed = true;
        }

        DisposeAll();
    }

    private string Register(IPresentationBatch batch)
    {
        lock (_lifecycleLock)
        {
            if (!_disposed)
            {
                var sessionId = Guid.NewGuid().ToString("N");
                _sessions[sessionId] = batch;
                return sessionId;
            }
        }

        // Startup may finish after shutdown has already drained the registered sessions.
        batch.DisposeForShutdown();
        throw new ObjectDisposedException(nameof(PresentationSessionRegistry));
    }

    /// <summary>
    /// Starts disposing <paramref name="batch"/> on the thread pool and tracks the task in
    /// <see cref="_pendingDisposals"/> until it completes, so <see cref="DisposeAll()"/> can find
    /// and await it later regardless of when it was started.
    /// </summary>
    private Task StartBackgroundDispose(string sessionId, IPresentationBatch batch, Task? pendingOperation = null, bool preserveUnsavedChanges = false)
    {
        var task = pendingOperation is null
            ? Task.Run(() => DisposeBatch(sessionId, batch, preserveUnsavedChanges))
            : pendingOperation.ContinueWith(
                _ => DisposeBatch(sessionId, batch, preserveUnsavedChanges),
                CancellationToken.None,
                TaskContinuationOptions.None,
                TaskScheduler.Default);
        _pendingDisposals[task] = 0;

        // Self-remove once finished so the tracking set doesn't grow unbounded over a long-lived
        // host process with many open/close cycles.
        task.ContinueWith(
            static (completed, state) =>
            {
                var registry = (PresentationSessionRegistry)state!;
                registry._pendingDisposals.TryRemove(completed, out _);
            },
            this,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);

        return task;
    }

    private void DisposeBatch(string sessionId, IPresentationBatch batch, bool preserveUnsavedChanges)
    {
        try
        {
            if (preserveUnsavedChanges)
                batch.DisposeForShutdown();
            else
                batch.Dispose();
        }
#pragma warning disable CA1031 // Shutdown cleanup must never throw — a failed dispose cannot block closing other sessions.
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to dispose PowerPoint session {SessionId}.", sessionId);
        }
#pragma warning restore CA1031
    }
}
