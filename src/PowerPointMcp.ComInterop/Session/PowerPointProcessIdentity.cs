namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Identifies a PowerPoint process by PID and creation time so PID reuse cannot transfer
/// ownership to an unrelated process.
/// </summary>
public readonly record struct PowerPointProcessIdentity(
    int ProcessId,
    long StartedAtUtcFileTime);
