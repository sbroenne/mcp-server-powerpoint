namespace Sbroenne.PowerPointMcp.Core.Notes;

/// <summary>
/// Result of a speaker notes operation (set/get notes text).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class NotesOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>The speaker notes text, for GetNotesText or after SetNotesText.</summary>
    public string? NotesText { get; init; }
}
