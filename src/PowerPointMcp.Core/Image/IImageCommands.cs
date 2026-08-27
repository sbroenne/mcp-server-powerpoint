using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Image commands: add a picture file to a slide. Operates within an already-open
/// IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
[ServiceCategory("image", "Image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "content",
    Description = "Insert embedded or linked pictures into slides and adjust picture appearance with brightness/contrast, recolor, and crop operations.")]
public interface IImageCommands
{
    /// <summary>
    /// Adds a picture from a local file to the given slide. The default embeds the picture
    /// (<paramref name="linkToFile"/> is false and <paramref name="saveWithDocument"/> is true).
    /// Set <paramref name="linkToFile"/> to true for a linked picture; set
    /// <paramref name="saveWithDocument"/> to false for a link-only picture that depends on the
    /// source path remaining available. The combination false/false is invalid.
    /// </summary>
    /// <param name="linkToFile">Whether the picture remains linked to its source file. Defaults to false.</param>
    /// <param name="saveWithDocument">Whether PowerPoint stores picture data in the presentation. Defaults to true.</param>
    ImageOperationResult AddPicture(
        IPresentationBatch batch,
        int slideIndex,
        string imagePath,
        float left,
        float top,
        float width,
        float height,
        bool linkToFile = false,
        bool saveWithDocument = true);

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
