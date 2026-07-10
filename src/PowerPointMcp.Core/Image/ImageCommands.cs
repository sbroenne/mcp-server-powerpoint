using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <inheritdoc cref="IImageCommands"/>
public sealed class ImageCommands : IImageCommands
{
    private const int MsoFalse = 0;
    private const int MsoTrue = -1;

    // MsoPictureColorType member name -> value, for SetRecolor/GetRecolor
    // (learn.microsoft.com/office/vba/api/office.msopicturecolortype) — verified live via
    // PictureEffectsDiagTests (a temporary diagnostic spike, since removed).
    private static readonly Dictionary<string, int> PictureColorTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoPictureAutomatic"] = 1,
        ["msoPictureGrayscale"] = 2,
        ["msoPictureBlackAndWhite"] = 3,
        ["msoPictureWatermark"] = 4,
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
            // — LinkToFile/SaveWithDocument are Microsoft.Office.Core.MsoTriState (office.dll)
            // typed, so called late-bound via dynamic with the raw int constants, same pattern
            // as elsewhere in this project. LinkToFile=False, SaveWithDocument=True embeds the
            // image directly in the .pptx rather than linking to the external file.
            dynamic slide = ctx.Presentation.Slides[slideIndex];
            slide.Shapes.AddPicture(fullImagePath, MsoFalse, MsoTrue, left, top, width, height);
            int newIndex = (int)slide.Shapes.Count; // always appended

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = (int)slide.Shapes.Count
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult SetBrightnessContrast(IPresentationBatch batch, int slideIndex, int shapeIndex, float brightness, float contrast)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic pictureFormat = shape.PictureFormat;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic pictureFormat = shape.PictureFormat;

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Brightness = (float)pictureFormat.Brightness,
                Contrast = (float)pictureFormat.Contrast,
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic pictureFormat = shape.PictureFormat;
            pictureFormat.ColorType = typeValue;

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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic pictureFormat = shape.PictureFormat;
            int typeValue = (int)pictureFormat.ColorType;
            string typeName = PictureColorTypesByValue.TryGetValue(typeValue, out var name) ? name : $"unknown({typeValue})";

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ColorTypeName = typeName,
            };
        });
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
