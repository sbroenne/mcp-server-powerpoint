using Sbroenne.PowerPointMcp.Core.Tags;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>
/// Result of a presentation lifecycle, template, metadata, or Mark as Final operation.
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant used throughout mcp-server-excel
/// (Rule 1): Success == true implies ErrorMessage is null/empty. Never set both.
/// </remarks>
public sealed class PresentationOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Full path to the presentation file the operation acted on.</summary>
    public string? PresentationPath { get; init; }

    /// <summary>
    /// The design/theme name currently applied to the presentation (set by
    /// <see cref="Sbroenne.PowerPointMcp.Core.Presentation.IPresentationCommands.ApplyTemplate"/>
    /// and <see cref="Sbroenne.PowerPointMcp.Core.Presentation.IPresentationCommands.GetThemeName"/>).
    /// Null for operations that don't touch theming.
    /// </summary>
    public string? ThemeName { get; init; }

    /// <summary>
    /// PowerPoint's Mark as Final state. This advisory editing flag is not authentication,
    /// encryption, or access control. Null for operations that do not read or update the flag.
    /// </summary>
    public bool? IsFinal { get; init; }

    /// <summary>
    /// The document property name acted on by the document-property/custom-property commands.
    /// Null for operations that don't touch document properties.
    /// </summary>
    public string? PropertyName { get; init; }

    /// <summary>
    /// The document property value read or written by the document-property/custom-property
    /// commands. Null for operations that don't touch document properties.
    /// </summary>
    public string? PropertyValue { get; init; }

    /// <summary>Normalized tag name for tag operations.</summary>
    public string? TagName { get; init; }

    /// <summary>Tag value for tag operations. Values are returned without case normalization.</summary>
    public string? TagValue { get; init; }

    /// <summary>1-based index of the tag in PowerPoint's native collection order.</summary>
    public int? TagIndex { get; init; }

    /// <summary>Number of string tags on the presentation.</summary>
    public int? TagCount { get; init; }

    /// <summary>Presentation tags in native 1-based collection order.</summary>
    public IReadOnlyList<TagInfo>? Tags { get; init; }
}
