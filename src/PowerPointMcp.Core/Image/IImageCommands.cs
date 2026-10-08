using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Image commands: add, adjust, crop, and compress pictures. Operates within an already-open
/// IPresentationBatch, targeting specific slides and shapes by their 1-based indexes.
/// </summary>
[ServiceCategory("image", "Image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "content",
    Description = "Insert embedded or linked pictures, adjust brightness/contrast and transparency, recolor and crop pictures, and compress presentation images.")]
[McpReadOnlyActions("get-brightness-contrast", "get-recolor", "get-crop", "get-transparency-color", "get-transparent-background", "get-crop-frame")]
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
        bool saveWithDocument = true,
        string compression = "default");

    /// <summary>Sets a picture shape's brightness and contrast (each 0-1, where 0.5 is PowerPoint's default/unadjusted level).</summary>
    ImageOperationResult SetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex, float brightness, float contrast);

    /// <summary>Gets a picture shape's current brightness and contrast (each 0-1).</summary>
    ImageOperationResult GetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Adjusts picture brightness by a relative amount. The amount is clamped by PowerPoint to the 0-1 range.</summary>
    ImageOperationResult IncrementBrightness(IPresentationBatch batch, int slideIndex, int shapeIndex, float increment);

    /// <summary>Adjusts picture contrast by a relative amount. The amount is clamped by PowerPoint to the 0-1 range.</summary>
    ImageOperationResult IncrementContrast(IPresentationBatch batch, int slideIndex, int shapeIndex, float increment);

    /// <summary>
    /// Recolors a picture shape. <paramref name="colorType"/> is an <c>MsoPictureColorType</c>
    /// enum member name: <c>"msoPictureAutomatic"</c> (original colors), <c>"msoPictureGrayscale"</c>,
    /// <c>"msoPictureBlackAndWhite"</c>, or <c>"msoPictureWatermark"</c> (washed-out, low-contrast).
    /// </summary>
    ImageOperationResult SetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex, string colorType);

    /// <summary>Gets a picture shape's current recolor mode as its <c>MsoPictureColorType</c> name.</summary>
    ImageOperationResult GetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the transparency color key as a PowerPoint RGB color integer (0xBBGGRR). The result
    /// applies only when transparent-background mode is enabled and can depend on the picture format.
    /// </summary>
    ImageOperationResult SetTransparencyColor(IPresentationBatch batch, int slideIndex, int shapeIndex, int colorRgb);

    /// <summary>
    /// Gets the picture's transparency color key as a PowerPoint RGB color integer (0xBBGGRR),
    /// or null when PowerPoint does not expose a valid 24-bit color key.
    /// </summary>
    ImageOperationResult GetTransparencyColor(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Enables or disables color-key transparency for a picture.</summary>
    ImageOperationResult SetTransparentBackground(IPresentationBatch batch, int slideIndex, int shapeIndex, bool enabled);

    /// <summary>Gets whether color-key transparency is enabled for a picture.</summary>
    ImageOperationResult GetTransparentBackground(IPresentationBatch batch, int slideIndex, int shapeIndex);

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

    /// <summary>
    /// Sets the full picture crop frame in points: source picture width/height and offsets within
    /// the frame, followed by the frame's position and size on the slide. Applies only to pictures.
    /// </summary>
    ImageOperationResult SetCropFrame(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        float pictureWidth,
        float pictureHeight,
        float pictureOffsetX,
        float pictureOffsetY,
        float frameLeft,
        float frameTop,
        float frameWidth,
        float frameHeight);

    /// <summary>Gets the source picture dimensions/offsets and crop-frame position/size in points.</summary>
    ImageOperationResult GetCropFrame(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Compresses one picture, all pictures on a slide, or all pictures in the presentation.
    /// Supply both indexes for one picture, only <paramref name="slideIndex"/> for a slide, or
    /// neither for the whole presentation. Resolution is <c>high-fidelity</c>, <c>hd</c> (330 PPI),
    /// <c>print</c> (220 PPI), <c>web</c> (150 PPI), or <c>email</c> (96 PPI).
    /// </summary>
    ImageOperationResult CompressPictures(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? shapeIndex = null,
        string resolution = "print",
        bool deleteCroppedAreas = false);
}
