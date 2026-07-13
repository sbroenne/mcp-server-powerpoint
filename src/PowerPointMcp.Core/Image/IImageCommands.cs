using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Image commands: embed a picture file into a slide. Operates within an already-open
/// IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
[ServiceCategory("image", "Image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "content",
    Description = "Insert pictures into slides and adjust picture appearance with brightness/contrast, recolor, and crop operations.")]
public interface IImageCommands
{
    /// <summary>
    /// Adds a picture from a local file to the given slide as an embedded shape (not linked to the
    /// source file).
    /// </summary>
    ImageOperationResult AddPicture(IPresentationBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height);

    /// <summary>Sets a picture shape's brightness and contrast (each 0-1, where 0.5 is PowerPoint's default/unadjusted level).</summary>
    ImageOperationResult SetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex, float brightness, float contrast);

    /// <summary>Gets a picture shape's current brightness and contrast (each 0-1).</summary>
    ImageOperationResult GetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Recolors a picture shape. <paramref name="colorType"/> is an <c>MsoPictureColorType</c>
    /// enum member name: <c>"msoPictureAutomatic"</c> (original colors), <c>"msoPictureGrayscale"</c>,
    /// <c>"msoPictureBlackAndWhite"</c>, or <c>"msoPictureWatermark"</c> (washed-out, low-contrast).
    /// </summary>
    ImageOperationResult SetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex, string colorType);

    /// <summary>Gets a picture shape's current recolor mode as its <c>MsoPictureColorType</c> name.</summary>
    ImageOperationResult GetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the crop offsets (in points) for all four sides of a picture shape.
    /// <paramref name="cropLeft"/>, <paramref name="cropTop"/>, <paramref name="cropRight"/>, and
    /// <paramref name="cropBottom"/> specify the amount to crop from each edge. Negative values are
    /// valid and expand the visible area beyond the image boundary; no clamping is applied.
    /// Units: points (1 pt = 1/72 inch). Applies to picture and linked-picture shapes only.
    /// </summary>
    ImageOperationResult SetCrop(IPresentationBatch batch, int slideIndex, int shapeIndex,
        float cropLeft, float cropTop, float cropRight, float cropBottom);

    /// <summary>
    /// Gets the current crop offsets (in points) for all four sides of a picture shape.
    /// Returns <see cref="ImageOperationResult.CropLeft"/>, <see cref="ImageOperationResult.CropTop"/>,
    /// <see cref="ImageOperationResult.CropRight"/>, and <see cref="ImageOperationResult.CropBottom"/>.
    /// Applies to picture and linked-picture shapes only.
    /// </summary>
    ImageOperationResult GetCrop(IPresentationBatch batch, int slideIndex, int shapeIndex);
}
