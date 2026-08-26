using Sbroenne.PowerPointMcp.Core.Tags;

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

    /// <summary>Second gradient stop color (0x00BBGGRR), for Set/GetGradientBackground.</summary>
    public int? ColorRgb2 { get; init; }

    /// <summary>MsoGradientStyle member name (e.g. "msoGradientHorizontal"), for Set/GetGradientBackground.</summary>
    public string? GradientStyleName { get; init; }

    /// <summary>Gradient variant (1-4), for Set/GetGradientBackground.</summary>
    public int? GradientVariant { get; init; }

    /// <summary>Whether the slide currently follows the slide master's background.</summary>
    public bool? FollowsMasterBackground { get; init; }

    /// <summary>1-based index of the section the operation created or acted on, if applicable.</summary>
    public int? SectionIndex { get; init; }

    /// <summary>Total section count in the presentation after the operation.</summary>
    public int? SectionCount { get; init; }

    /// <summary>Name of the section, if applicable.</summary>
    public string? SectionName { get; init; }

    /// <summary>Legacy comments on a slide, in 1-based comment order.</summary>
    public IReadOnlyList<SlideCommentInfo>? Comments { get; init; }

    /// <summary>Number of legacy comments on the slide.</summary>
    public int? CommentCount { get; init; }

    /// <summary>Number of slides imported by an import operation.</summary>
    public int? ImportedSlideCount { get; init; }

    /// <summary>New 1-based indexes assigned to imported slides.</summary>
    public IReadOnlyList<int>? ImportedSlideIndexes { get; init; }

    /// <summary>Normalized tag name for tag operations.</summary>
    public string? TagName { get; init; }

    /// <summary>Tag value for tag operations. Values are returned without case normalization.</summary>
    public string? TagValue { get; init; }

    /// <summary>1-based index of the tag in PowerPoint's native collection order.</summary>
    public int? TagIndex { get; init; }

    /// <summary>Number of string tags on the slide.</summary>
    public int? TagCount { get; init; }

    /// <summary>Slide tags in native 1-based collection order.</summary>
    public IReadOnlyList<TagInfo>? Tags { get; init; }
}
