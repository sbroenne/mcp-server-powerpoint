using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Session;

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
    bool IsPowerPointProcessAlive);

/// <summary>
/// In-process registry that maps a generated session identifier to a long-lived
/// <see cref="IPresentationBatch"/>. Registered as a DI singleton for the lifetime of the
/// MCP host so that one PowerPoint session can span many tool invocations.
/// </summary>
/// <remarks>
/// This is the MVP replacement for mcp-server-excel's out-of-process Service + named-pipe
/// bridge (see .squad/decisions.md — Dallas's McpServer architecture pass). Because the STA
/// thread and work queue already live inside <c>PresentationBatch</c>, an in-process
/// dictionary fully satisfies "one long-lived session across many tool invocations" without
/// RPC. Disposing a batch closes PowerPoint for that presentation, so the host MUST call
/// <see cref="DisposeAll()"/> on shutdown to guarantee no lingering POWERPNT.exe process.
///
/// ASYNC CLOSE (2026-07-01, see .squad/decisions/inbox/brett-async-close.md): Parker's shutdown
/// hardening made <c>IPresentationBatch.Dispose()</c> legitimately block for up to
/// <see cref="ComInteropConstants.StaThreadJoinTimeout"/> (~210s) in the worst case — the bounded
/// grace period + force-kill safety net that guarantees no orphaned POWERPNT.exe. An MCP tool
/// call cannot block that long without risking a client-side timeout, so <see cref="Close"/>
/// atomically removes the session and starts its dispose on a background <see cref="Task"/>,
/// returning immediately. In-flight dispose tasks are tracked so <see cref="DisposeAll()"/> (host
/// shutdown) can await every one of them — both the ones it starts itself and any still pending
/// from earlier <see cref="Close"/> calls — before the host is allowed to exit. This preserves
/// the "no lingering POWERPNT.exe after shutdown" guarantee while keeping the MCP-visible
/// `close_presentation` call fast.
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

    private readonly ConcurrentDictionary<string, IPresentationBatch> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<Task, byte> _pendingDisposals = new();
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
    /// (closing PowerPoint for that presentation) on a background task. Returns immediately —
    /// does NOT wait for PowerPoint's Quit/grace-period/force-kill sequence to finish, which can
    /// legitimately take up to <see cref="ComInteropConstants.StaThreadJoinTimeout"/>. The
    /// background dispose is tracked so <see cref="DisposeAll()"/> can await it on host shutdown.
    /// No-op if the identifier is unknown.
    /// </summary>
    /// <param name="sessionId">The session identifier to close.</param>
    /// <returns><see langword="true"/> if a session was found and its close was started.</returns>
    public bool Close(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId) || !_sessions.TryRemove(sessionId, out var batch))
        {
            return false;
        }

        StartBackgroundDispose(sessionId, batch);
        return true;
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
                pair.Value.IsPowerPointProcessAlive()));
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
        foreach (var sessionId in _sessions.Keys.ToArray())
        {
            if (_sessions.TryRemove(sessionId, out var batch))
            {
                StartBackgroundDispose(sessionId, batch);
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
                    "Timed out after {TimeoutSeconds:N0}s waiting for {Count} PowerPoint session(s) to finish shutting down in the background; each session's own force-kill safety net remains active and will still run to completion.",
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
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        DisposeAll();
    }

    private string Register(IPresentationBatch batch)
    {
        var sessionId = Guid.NewGuid().ToString("N");
        _sessions[sessionId] = batch;
        return sessionId;
    }

    /// <summary>
    /// Starts disposing <paramref name="batch"/> on the thread pool and tracks the task in
    /// <see cref="_pendingDisposals"/> until it completes, so <see cref="DisposeAll()"/> can find
    /// and await it later regardless of when it was started.
    /// </summary>
    private Task StartBackgroundDispose(string sessionId, IPresentationBatch batch)
    {
        var task = Task.Run(() => DisposeBatch(sessionId, batch));
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

    private void DisposeBatch(string sessionId, IPresentationBatch batch)
    {
        try
        {
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
