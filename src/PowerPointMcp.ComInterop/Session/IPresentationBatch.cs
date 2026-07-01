namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Represents a batch of PowerPoint operations that share a single PowerPoint instance.
/// Implements IDisposable to ensure proper COM cleanup.
/// </summary>
/// <remarks>
/// Use this interface via PresentationSession.BeginBatch() for multi-operation workflows.
/// The batch keeps PowerPoint and the presentation open until disposed, enabling efficient
/// execution of multiple commands without repeated PowerPoint startup/shutdown overhead.
///
/// This mirrors mcp-server-excel's IExcelBatch pattern (STA thread, channel-based work
/// queue, OLE message filter, operation timeout). Current scope: single presentation only
/// (no multi-presentation cross-file operations, no IRM detection) — see plan/continuation
/// notes for follow-up work porting those Excel-specific hardening features.
/// </remarks>
public interface IPresentationBatch : IDisposable
{
    /// <summary>Gets the path to the PowerPoint presentation this batch operates on.</summary>
    string PresentationPath { get; }

    /// <summary>
    /// Executes a void COM operation within this batch. The operation receives a
    /// PresentationContext with access to the PowerPoint app and presentation.
    /// All PowerPoint COM operations are synchronous.
    /// </summary>
    void Execute(
        Action<PresentationContext, CancellationToken> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a COM operation within this batch and returns its result.
    /// </summary>
    T Execute<T>(
        Func<PresentationContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves changes to the presentation. This is an explicit save — changes are NOT
    /// automatically saved on dispose.
    /// </summary>
    void Save(CancellationToken cancellationToken = default);

    /// <summary>Checks if the underlying PowerPoint process is still alive.</summary>
    bool IsPowerPointProcessAlive();

    /// <summary>
    /// Gets whether a previous operation timed out or was cancelled while PowerPoint was
    /// unresponsive. When true, the session is poisoned and callers must fail fast.
    /// </summary>
    bool HasTimedOutOperation { get; }

    /// <summary>Gets the PowerPoint process ID, if captured.</summary>
    int? PowerPointProcessId { get; }

    /// <summary>Gets the operation timeout for this batch.</summary>
    TimeSpan OperationTimeout { get; }
}
