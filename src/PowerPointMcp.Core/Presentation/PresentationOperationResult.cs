namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>
/// Result of a presentation lifecycle operation (open, create, close, save).
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
    /// The document property name acted on by the document-property/custom-property commands.
    /// Null for operations that don't touch document properties.
    /// </summary>
    public string? PropertyName { get; init; }

    /// <summary>
    /// The document property value read or written by the document-property/custom-property
    /// commands. Null for operations that don't touch document properties.
    /// </summary>
    public string? PropertyValue { get; init; }
}
