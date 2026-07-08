namespace Sbroenne.PowerPointMcp.Core.Master;

/// <summary>
/// Result of a slide master operation (title/body placeholder font, master background color).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class MasterOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Font name of the master title/body placeholder, for Get*Font or after Set*Font.</summary>
    public string? FontName { get; init; }

    /// <summary>Font size in points of the master title/body placeholder, if applicable.</summary>
    public float? FontSize { get; init; }

    /// <summary>Whether the master title/body placeholder font is bold, if applicable.</summary>
    public bool? Bold { get; init; }

    /// <summary>Font or background color as an RGB integer (0xBBGGRR, PowerPoint's native color order), if applicable.</summary>
    public int? ColorRgb { get; init; }
}
