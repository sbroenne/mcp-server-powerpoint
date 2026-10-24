namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>Exposes a returned operation's outcome without depending on Core result types.</summary>
public interface IPresentationOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    bool Success { get; }

    /// <summary>Error details for a failed operation; null on success.</summary>
    string? ErrorMessage { get; }
}
