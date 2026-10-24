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

    /// <summary>
    /// Saves a temporary copy of the current Open XML presentation, lets a non-COM transformation
    /// update that copy, then safely replaces and reopens the active file. The transformation runs
    /// while the original presentation remains open and must not access PowerPoint COM objects.
    /// </summary>
    T TransformPresentationCopy<T>(
        Func<string, CancellationToken, T> transform,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Updates the tracked presentation path after PowerPoint Save As changes the active file
    /// identity. Must be called from inside this batch's serialized <see cref="Execute(Action{PresentationContext, CancellationToken}, CancellationToken)"/>
    /// or <see cref="Execute{T}(Func{PresentationContext, CancellationToken, T}, CancellationToken)"/>
    /// callback immediately after the COM operation succeeds.
    /// </summary>
    /// <param name="presentationPath">New absolute presentation path.</param>
    void UpdatePresentationPath(string presentationPath);

    /// <summary>Checks if the underlying PowerPoint process is still alive.</summary>
    bool IsPowerPointProcessAlive();

    /// <summary>
    /// Gets whether an expired wait still has unfinished PowerPoint work.
    /// </summary>
    bool HasTimedOutOperation { get; }

    /// <summary>Reads execution state without waiting for the PowerPoint thread.</summary>
    PresentationOperationStatus GetOperationStatus();

    /// <summary>Tracks a complete service request, including waits between PowerPoint callbacks, and rejects other callers until it completes.</summary>
    void BeginRequest();

    /// <summary>Records a request outcome without hiding an unfinished PowerPoint callback.</summary>
    void CompleteRequest(string state, string? errorMessage = null);

    /// <summary>Records that a caller stopped waiting without stopping PowerPoint work.</summary>
    void MarkWaitExpired();

    /// <summary>Releases automation on host shutdown, leaving unsaved presentations open.</summary>
    void DisposeForShutdown();

    /// <summary>Gets the PowerPoint process ID, if captured.</summary>
    int? PowerPointProcessId { get; }

    /// <summary>Gets the PID plus creation time of the owned PowerPoint process, if captured.</summary>
    PowerPointProcessIdentity? PowerPointProcessIdentity { get; }

    /// <summary>Gets the operation timeout for this batch.</summary>
    TimeSpan OperationTimeout { get; }
}
