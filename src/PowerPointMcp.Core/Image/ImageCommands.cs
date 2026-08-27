extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <inheritdoc cref="IImageCommands"/>
public sealed class ImageCommands : IImageCommands
{
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
    public ImageOperationResult AddPicture(
        IPresentationBatch batch,
        int slideIndex,
        string imagePath,
        float left,
        float top,
        float width,
        float height,
        bool linkToFile = false,
        bool saveWithDocument = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(imagePath);

        if (!linkToFile && !saveWithDocument)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = "saveWithDocument must be true when linkToFile is false; otherwise PowerPoint has no picture data to retain."
            };
        }

        var pathValidation = ResolveImagePath(imagePath);
        if (pathValidation.Error is not null)
        {
            return pathValidation.Error;
        }

        string fullImagePath = pathValidation.FullPath!;
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
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? picture = null;
            try
            {
                slides = ctx.Presentation.Slides;
                int slideCount = slides.Count;
                if (slideIndex < 1 || slideIndex > slideCount)
                {
                    return new ImageOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                    };
                }

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                picture = shapes.AddPicture(
                    fullImagePath,
                    linkToFile ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                    saveWithDocument ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                    left,
                    top,
                    width,
                    height);
                int shapeCount = shapes.Count;

                return new ImageOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeCount,
                    ShapeCount = shapeCount,
                    LinkToFile = linkToFile,
                    SaveWithDocument = saveWithDocument
                };
            }
            finally
            {
                ComUtilities.Release(ref picture);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
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

            PowerPoint.PictureFormat? pictureFormat = null;
            try
            {
                pictureFormat = shape.PictureFormat;
                pictureFormat.Brightness = brightness;
                pictureFormat.Contrast = contrast;

                return new ImageOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Brightness = brightness,
                    Contrast = contrast,
                };
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }
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

            PowerPoint.PictureFormat? pictureFormat = null;
            try
            {
                pictureFormat = shape.PictureFormat;

                return new ImageOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Brightness = pictureFormat.Brightness,
                    Contrast = pictureFormat.Contrast,
                };
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }
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
            dynamic? pictureFormat = null;
            try
            {
                pictureFormat = shape.PictureFormat;
                pictureFormat.ColorType = typeValue;
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }

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
            dynamic? pictureFormat = null;
            int rawColorType;
            try
            {
                pictureFormat = shape.PictureFormat;
                rawColorType = (int)pictureFormat.ColorType;
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }
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
            PowerPoint.PictureFormat? pictureFormat = null;
            try
            {
                pictureFormat = shape.PictureFormat;
                pictureFormat.CropLeft = cropLeft;
                pictureFormat.CropTop = cropTop;
                pictureFormat.CropRight = cropRight;
                pictureFormat.CropBottom = cropBottom;

                return new ImageOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    CropLeft = pictureFormat.CropLeft,
                    CropTop = pictureFormat.CropTop,
                    CropRight = pictureFormat.CropRight,
                    CropBottom = pictureFormat.CropBottom,
                };
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }
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

            PowerPoint.PictureFormat? pictureFormat = null;
            try
            {
                pictureFormat = shape.PictureFormat;

                return new ImageOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    CropLeft = pictureFormat.CropLeft,
                    CropTop = pictureFormat.CropTop,
                    CropRight = pictureFormat.CropRight,
                    CropBottom = pictureFormat.CropBottom,
                };
            }
            finally
            {
                if (pictureFormat != null)
                {
                    ComUtilities.Release(ref pictureFormat!);
                }
            }
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
    /// accessing <c>PictureFormat</c> members).
    /// </summary>
    private static ImageOperationResult? ValidatePictureShape(PowerPoint.Shape shape, int slideIndex, int shapeIndex)
    {
        Office.MsoShapeType shapeType = shape.Type;
        if (shapeType is not Office.MsoShapeType.msoPicture and not Office.MsoShapeType.msoLinkedPicture)
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a picture (shape type={shapeType}). PictureFormat operations require a picture or linked picture shape."
            };
        }
        return null;
    }

    private static (string? FullPath, ImageOperationResult? Error) ResolveImagePath(string imagePath)
    {
        if (!Path.IsPathFullyQualified(imagePath))
        {
            return (null, new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Image path must be a full local or UNC path: {imagePath}."
            });
        }

        try
        {
            return (Path.GetFullPath(imagePath), null);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return (null, new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Image path cannot be resolved: {ex.Message}"
            });
        }
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
