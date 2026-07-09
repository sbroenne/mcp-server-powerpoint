namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>
/// Result of a slide operation (add, delete, count).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as PresentationOperationResult
/// (Rule 1): Success == true implies ErrorMessage is null/empty.
/// </remarks>
public sealed class SlideOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the slide the operation created or acted on, if applicable.</summary>
    public int? SlideIndex { get; init; }

    /// <summary>Total slide count in the presentation after the operation.</summary>
    public int? SlideCount { get; init; }

    /// <summary>Background solid fill color, packed as 0x00BBGGRR (matches VBA's RGB()).</summary>
    public int? ColorRgb { get; init; }

    /// <summary>Whether the slide currently follows the slide master's background.</summary>
    public bool? FollowsMasterBackground { get; init; }

    /// <summary>1-based index of the section the operation created or acted on, if applicable.</summary>
    public int? SectionIndex { get; init; }

    /// <summary>Total section count in the presentation after the operation.</summary>
    public int? SectionCount { get; init; }

    /// <summary>Name of the section, if applicable.</summary>
    public string? SectionName { get; init; }
}
