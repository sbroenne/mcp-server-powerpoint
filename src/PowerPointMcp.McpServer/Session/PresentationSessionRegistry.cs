using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
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
/// <see cref="DisposeAll"/> on shutdown to guarantee no lingering POWERPNT.exe process.
/// </remarks>
public sealed class PresentationSessionRegistry : IDisposable
{
    private readonly ConcurrentDictionary<string, IPresentationBatch> _sessions = new(StringComparer.Ordinal);
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
    /// Closes a session: removes it from the registry and disposes its batch (closing PowerPoint
    /// for that presentation). No-op if the identifier is unknown.
    /// </summary>
    /// <param name="sessionId">The session identifier to close.</param>
    /// <returns><see langword="true"/> if a session was found and closed.</returns>
    public bool Close(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId) || !_sessions.TryRemove(sessionId, out var batch))
        {
            return false;
        }

        DisposeBatch(sessionId, batch);
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
    /// Disposes every registered batch and clears the registry. Idempotent and safe to call
    /// from host shutdown as well as from <see cref="Dispose"/>.
    /// </summary>
    public void DisposeAll()
    {
        foreach (var sessionId in _sessions.Keys.ToArray())
        {
            if (_sessions.TryRemove(sessionId, out var batch))
            {
                DisposeBatch(sessionId, batch);
            }
        }
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
