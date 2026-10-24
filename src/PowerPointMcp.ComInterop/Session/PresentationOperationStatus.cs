namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>Execution metadata that can be read without calling PowerPoint.</summary>
public sealed record PresentationOperationStatus(
    string State,
    bool ExecutionPending,
    bool WaitExpired,
    DateTime? StartedAtUtc,
    string? ErrorMessage);
