namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Result of an image operation.
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class ImageOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the picture shape that was added.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Total shape count on the slide after the operation.</summary>
    public int? ShapeCount { get; init; }

    /// <summary>Picture brightness (0-1), if applicable.</summary>
    public float? Brightness { get; init; }

    /// <summary>Picture contrast (0-1), if applicable.</summary>
    public float? Contrast { get; init; }

    /// <summary>The MsoPictureColorType name of the picture's recolor mode, if applicable.</summary>
    public string? ColorTypeName { get; init; }

    /// <summary>Crop offset for the left edge of the picture, in points. Null when not applicable.</summary>
    public float? CropLeft { get; init; }

    /// <summary>Crop offset for the top edge of the picture, in points. Null when not applicable.</summary>
    public float? CropTop { get; init; }

    /// <summary>Crop offset for the right edge of the picture, in points. Null when not applicable.</summary>
    public float? CropRight { get; init; }

    /// <summary>Crop offset for the bottom edge of the picture, in points. Null when not applicable.</summary>
    public float? CropBottom { get; init; }
}
