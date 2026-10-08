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

    /// <summary>Whether the added picture links to its source file, if applicable.</summary>
    public bool? LinkToFile { get; init; }

    /// <summary>Whether a copy of the added picture is saved in the presentation, if applicable.</summary>
    public bool? SaveWithDocument { get; init; }

    /// <summary>The picture compression mode used on insertion, when applicable.</summary>
    public string? CompressionMode { get; init; }

    /// <summary>Picture brightness (0-1), if applicable.</summary>
    public float? Brightness { get; init; }

    /// <summary>Picture contrast (0-1), if applicable.</summary>
    public float? Contrast { get; init; }

    /// <summary>The MsoPictureColorType name of the picture's recolor mode, if applicable.</summary>
    public string? ColorTypeName { get; init; }

    /// <summary>Transparency color key as a 24-bit RGB integer (0xBBGGRR), if applicable.</summary>
    public int? ColorRgb { get; init; }

    /// <summary>Whether color-key transparency is enabled, if applicable.</summary>
    public bool? TransparentBackground { get; init; }

    /// <summary>Crop offset for the left edge of the picture, in points. Null when not applicable.</summary>
    public float? CropLeft { get; init; }

    /// <summary>Crop offset for the top edge of the picture, in points. Null when not applicable.</summary>
    public float? CropTop { get; init; }

    /// <summary>Crop offset for the right edge of the picture, in points. Null when not applicable.</summary>
    public float? CropRight { get; init; }

    /// <summary>Crop offset for the bottom edge of the picture, in points. Null when not applicable.</summary>
    public float? CropBottom { get; init; }

    /// <summary>Original source picture width inside the crop frame, in points.</summary>
    public float? PictureWidth { get; init; }

    /// <summary>Original source picture height inside the crop frame, in points.</summary>
    public float? PictureHeight { get; init; }

    /// <summary>Horizontal source picture offset within the crop frame, in points.</summary>
    public float? PictureOffsetX { get; init; }

    /// <summary>Vertical source picture offset within the crop frame, in points.</summary>
    public float? PictureOffsetY { get; init; }

    /// <summary>Left position of the crop frame on the slide, in points.</summary>
    public float? FrameLeft { get; init; }

    /// <summary>Top position of the crop frame on the slide, in points.</summary>
    public float? FrameTop { get; init; }

    /// <summary>Width of the crop frame, in points.</summary>
    public float? FrameWidth { get; init; }

    /// <summary>Height of the crop frame, in points.</summary>
    public float? FrameHeight { get; init; }

    /// <summary>Requested compression resolution preset.</summary>
    public string? Resolution { get; init; }

    /// <summary>Number of pictures recompressed.</summary>
    public int? CompressedPictureCount { get; init; }

    /// <summary>Number of pictures skipped because they are linked or use unsupported formats.</summary>
    public int? SkippedPictureCount { get; init; }

    /// <summary>Total bytes in processed image parts before recompression.</summary>
    public long? OriginalImageBytes { get; init; }

    /// <summary>Total bytes in processed image parts after recompression.</summary>
    public long? CompressedImageBytes { get; init; }

    /// <summary>Picture descriptions for images that could not be recompressed.</summary>
    public IReadOnlyList<string>? SkippedPictures { get; init; }
}
