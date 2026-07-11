namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>
/// Result of a text frame operation (set/get text, font formatting).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class TextFrameOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>The text content, for GetText or after SetText.</summary>
    public string? Text { get; init; }

    /// <summary>Font size in points, if applicable.</summary>
    public float? FontSize { get; init; }

    /// <summary>Whether the font is bold, if applicable.</summary>
    public bool? Bold { get; init; }

    /// <summary>Font color as an RGB integer (0xBBGGRR, PowerPoint's native color order), if applicable.</summary>
    public int? ColorRgb { get; init; }

    /// <summary>Whether the font is italic, if applicable.</summary>
    public bool? Italic { get; init; }

    /// <summary>Whether the font is underlined, if applicable.</summary>
    public bool? Underline { get; init; }

    /// <summary>Font name (typeface), if applicable.</summary>
    public string? FontName { get; init; }

    /// <summary>The PpParagraphAlignment name of the text range's paragraph alignment, if applicable.</summary>
    public string? Alignment { get; init; }

    /// <summary>Whether bullets are enabled for the text range, if applicable.</summary>
    public bool? BulletEnabled { get; init; }

    /// <summary>The bullet glyph character, if bullets are enabled and applicable.</summary>
    public string? BulletCharacter { get; init; }

    /// <summary>The PpAutoSize name of the text frame's auto-size mode, if applicable.</summary>
    public string? AutoSize { get; init; }
}
