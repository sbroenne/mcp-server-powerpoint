using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <inheritdoc cref="IImageCommands"/>
public sealed class ImageCommands : IImageCommands
{
    // MsoTriState values from Microsoft.Office.Core — office.dll is not referenced/embedded,
    // so passed as raw ints via dynamic late binding (same pattern as ShapeCommands.cs).
    private const int MsoFalse = 0;   // MsoTriState.msoFalse
    private const int MsoTrue  = -1;  // MsoTriState.msoTrue

    // MsoShapeType values for picture-shape validation.
    // Shape.Type is Microsoft.Office.Core.MsoShapeType (Office.Core — not embedded);
    // read via (int)((dynamic)shape).Type and compared against these named constants.
    private const int MsoPicture       = 13; // MsoShapeType.msoPicture
    private const int MsoLinkedPicture = 11; // MsoShapeType.msoLinkedPicture

    // MsoPictureColorType member name -> value, for SetRecolor/GetRecolor
    // (learn.microsoft.com/office/vba/api/office.msopicturecolortype) — verified live via
    // PictureEffectsDiagTests (a temporary diagnostic spike, since removed).
    private static readonly Dictionary<string, int> PictureColorTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoPictureAutomatic"]     = 1,
        ["msoPictureGrayscale"]     = 2,
        ["msoPictureBlackAndWhite"] = 3,
        ["msoPictureWatermark"]     = 4,
    };

    private static readonly Dictionary<int, string> PictureColorTypesByValue =
        PictureColorTypes.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    /// <inheritdoc/>
    public ImageOperationResult AddPicture(IPresentationBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(imagePath);

        // Up-front validation (not a suppressed exception, Rule 1b) — gives a clear error
        // message instead of a generic COMException from AddPicture.
        string fullImagePath = Path.GetFullPath(imagePath);
        if (!File.Exists(fullImagePath))
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Image file not found: {fullImagePath}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new ImageOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            // Shapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            // — LinkToFile/SaveWithDocument are Microsoft.Office.Core.MsoTriState (office.dll),
            // so called late-bound via dynamic with the raw int constants (same pattern as
            // ShapeCommands.cs). LinkToFile=False, SaveWithDocument=True embeds the image
            // directly in the .pptx rather than linking to the external file.
            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            // Reason: Shapes.AddPicture is called late-bound via dynamic because
            // LinkToFile/SaveWithDocument are Microsoft.Office.Core.MsoTriState (office.dll).
            ((dynamic)slide.Shapes).AddPicture(fullImagePath, MsoFalse, MsoTrue, left, top, width, height);
            int newIndex = slide.Shapes.Count; // always appended

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult SetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex, float brightness, float contrast)
    {
        ArgumentNullException.ThrowIfNull(batch);

        // Pre-COM range validation (Rule 1b: checked before touching COM, not catch-and-return).
        var rangeValidation = ValidateBrightnessContrastRange(brightness, contrast);
        if (rangeValidation is not null) return rangeValidation;

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            PowerPoint.PictureFormat pictureFormat = shape.PictureFormat;
            pictureFormat.Brightness = brightness;
            pictureFormat.Contrast = contrast;

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Brightness = brightness,
                Contrast = contrast,
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult GetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            PowerPoint.PictureFormat pictureFormat = shape.PictureFormat;

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Brightness = pictureFormat.Brightness,
                Contrast = pictureFormat.Contrast,
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult SetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex, string colorType)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(colorType);

        if (!PictureColorTypes.TryGetValue(colorType, out var typeValue))
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"'{colorType}' is not a recognized MsoPictureColorType member name (must be 'msoPictureAutomatic', 'msoPictureGrayscale', 'msoPictureBlackAndWhite', or 'msoPictureWatermark')."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            // Reason: PictureFormat.ColorType is MsoPictureColorType (Microsoft.Office.Core — not embedded);
            // assigned via dynamic late binding with the pre-validated integer value.
            ((dynamic)shape.PictureFormat).ColorType = typeValue;

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ColorTypeName = colorType,
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult GetRecolor(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            // Reason: PictureFormat.ColorType is MsoPictureColorType (Microsoft.Office.Core — not embedded);
            // read via dynamic late binding.
            int rawColorType = (int)((dynamic)shape.PictureFormat).ColorType;
            string typeName = PictureColorTypesByValue.TryGetValue(rawColorType, out var name) ? name : $"unknown({rawColorType})";

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ColorTypeName = typeName,
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult SetCrop(IPresentationBatch batch, int slideIndex, int shapeIndex,
        float cropLeft, float cropTop, float cropRight, float cropBottom)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            // CropLeft/Top/Right/Bottom are typed float properties on the embedded PIA.
            // Negative values are valid (expand visible area beyond image boundary); no clamping.
            PowerPoint.PictureFormat pictureFormat = shape.PictureFormat;
            pictureFormat.CropLeft   = cropLeft;
            pictureFormat.CropTop    = cropTop;
            pictureFormat.CropRight  = cropRight;
            pictureFormat.CropBottom = cropBottom;

            return new ImageOperationResult
            {
                Success    = true,
                ShapeIndex = shapeIndex,
                CropLeft   = pictureFormat.CropLeft,
                CropTop    = pictureFormat.CropTop,
                CropRight  = pictureFormat.CropRight,
                CropBottom = pictureFormat.CropBottom,
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult GetCrop(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
            if (typeValidation is not null) return typeValidation;

            PowerPoint.PictureFormat pictureFormat = shape.PictureFormat;

            return new ImageOperationResult
            {
                Success    = true,
                ShapeIndex = shapeIndex,
                CropLeft   = pictureFormat.CropLeft,
                CropTop    = pictureFormat.CropTop,
                CropRight  = pictureFormat.CropRight,
                CropBottom = pictureFormat.CropBottom,
            };
        });
    }

    /// <summary>
    /// Validates that brightness and contrast are each in [0, 1].
    /// Called before <c>batch.Execute</c> so range errors are caught without touching COM.
    /// </summary>
    private static ImageOperationResult? ValidateBrightnessContrastRange(float brightness, float contrast)
    {
        if (brightness < 0f || brightness > 1f)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Brightness {brightness} is out of range; must be between 0 and 1 (inclusive)."
            };
        }
        if (contrast < 0f || contrast > 1f)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Contrast {contrast} is out of range; must be between 0 and 1 (inclusive)."
            };
        }
        return null;
    }

    /// <summary>
    /// Validates that <paramref name="shape"/> is a picture or linked picture (required before
    /// accessing <c>PictureFormat</c> members). <c>Shape.Type</c> is
    /// <c>Microsoft.Office.Core.MsoShapeType</c> (Office.Core — not embedded), so the check
    /// uses dynamic late binding with named integer constants.
    /// </summary>
    private static ImageOperationResult? ValidatePictureShape(PowerPoint.Shape shape, int slideIndex, int shapeIndex)
    {
        // Reason: Shape.Type is Microsoft.Office.Core.MsoShapeType (Office.Core — not embedded),
        // so it is read via dynamic late binding and compared against named integer constants.
        int shapeType = (int)((dynamic)shape).Type;
        if (shapeType != MsoPicture && shapeType != MsoLinkedPicture)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a picture (shape type={shapeType}). PictureFormat operations require a picture or linked picture shape."
            };
        }
        return null;
    }

    private static ImageOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }

    private static ImageOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }
        return null;
    }
}
