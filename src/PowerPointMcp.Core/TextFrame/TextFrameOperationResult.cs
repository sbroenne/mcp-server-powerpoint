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
}
