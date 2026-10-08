extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Image;

public sealed partial class ImageCommands
{
    /// <inheritdoc/>
    public ImageOperationResult IncrementBrightness(
        IPresentationBatch batch, int slideIndex, int shapeIndex, float increment)
    {
        if (!float.IsFinite(increment))
        {
            return Failure("Brightness increment must be a finite number.");
        }

        return WithPicture(batch, slideIndex, shapeIndex, pictureFormat =>
        {
            pictureFormat.IncrementBrightness(increment);
            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Brightness = pictureFormat.Brightness
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult IncrementContrast(
        IPresentationBatch batch, int slideIndex, int shapeIndex, float increment)
    {
        if (!float.IsFinite(increment))
        {
            return Failure("Contrast increment must be a finite number.");
        }

        return WithPicture(batch, slideIndex, shapeIndex, pictureFormat =>
        {
            pictureFormat.IncrementContrast(increment);
            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Contrast = pictureFormat.Contrast
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult SetTransparencyColor(
        IPresentationBatch batch, int slideIndex, int shapeIndex, int colorRgb)
    {
        if (colorRgb is < 0 or > 0xFFFFFF)
        {
            return Failure("colorRgb must be between 0 and 16777215 (0x000000-0xFFFFFF).");
        }

        return WithPicture(batch, slideIndex, shapeIndex, pictureFormat =>
        {
            pictureFormat.TransparencyColor = colorRgb;
            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ColorRgb = NormalizeTransparencyColor(pictureFormat.TransparencyColor)
            };
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult GetTransparencyColor(
        IPresentationBatch batch, int slideIndex, int shapeIndex) =>
        WithPicture(batch, slideIndex, shapeIndex, pictureFormat => new ImageOperationResult
        {
            Success = true,
            ShapeIndex = shapeIndex,
            ColorRgb = NormalizeTransparencyColor(pictureFormat.TransparencyColor)
        });

    internal static int? NormalizeTransparencyColor(int colorRgb) =>
        colorRgb is >= 0 and <= 0xFFFFFF ? colorRgb : null;

    /// <inheritdoc/>
    public ImageOperationResult SetTransparentBackground(
        IPresentationBatch batch, int slideIndex, int shapeIndex, bool enabled) =>
        WithPicture(batch, slideIndex, shapeIndex, pictureFormat =>
        {
            pictureFormat.TransparentBackground = enabled
                ? Office.MsoTriState.msoTrue
                : Office.MsoTriState.msoFalse;
            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                TransparentBackground = pictureFormat.TransparentBackground == Office.MsoTriState.msoTrue
            };
        });

    /// <inheritdoc/>
    public ImageOperationResult GetTransparentBackground(
        IPresentationBatch batch, int slideIndex, int shapeIndex) =>
        WithPicture(batch, slideIndex, shapeIndex, pictureFormat => new ImageOperationResult
        {
            Success = true,
            ShapeIndex = shapeIndex,
            TransparentBackground = pictureFormat.TransparentBackground == Office.MsoTriState.msoTrue
        });

    /// <inheritdoc/>
    public ImageOperationResult SetCropFrame(
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
        float frameHeight)
    {
        if (!float.IsFinite(pictureWidth) || pictureWidth <= 0 ||
            !float.IsFinite(pictureHeight) || pictureHeight <= 0 ||
            !float.IsFinite(pictureOffsetX) || !float.IsFinite(pictureOffsetY) ||
            !float.IsFinite(frameLeft) || !float.IsFinite(frameTop) ||
            !float.IsFinite(frameWidth) || frameWidth <= 0 ||
            !float.IsFinite(frameHeight) || frameHeight <= 0)
        {
            return Failure("Picture and frame dimensions must be positive finite numbers; positions and offsets must be finite.");
        }

        return WithCrop(batch, slideIndex, shapeIndex, crop =>
        {
            crop.ShapeLeft = frameLeft;
            crop.ShapeTop = frameTop;
            crop.ShapeWidth = frameWidth;
            crop.ShapeHeight = frameHeight;
            crop.PictureWidth = pictureWidth;
            crop.PictureHeight = pictureHeight;
            crop.PictureOffsetX = pictureOffsetX;
            crop.PictureOffsetY = pictureOffsetY;
            return ReadCropFrame(shapeIndex, crop);
        });
    }

    /// <inheritdoc/>
    public ImageOperationResult GetCropFrame(
        IPresentationBatch batch, int slideIndex, int shapeIndex) =>
        WithCrop(batch, slideIndex, shapeIndex, crop => ReadCropFrame(shapeIndex, crop));

    private static ImageOperationResult WithPicture(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        Func<PowerPoint.PictureFormat, ImageOperationResult> operation)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.PictureFormat? pictureFormat = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null) return slideValidation;

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                var shapeValidation = ValidateShapeIndex(shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = shapes[shapeIndex];
                var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
                if (typeValidation is not null) return typeValidation;

                pictureFormat = shape.PictureFormat;
                return operation(pictureFormat);
            }
            finally
            {
                ComUtilities.Release(ref pictureFormat);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
        });
    }

    private static ImageOperationResult WithCrop(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        Func<Office.Crop, ImageOperationResult> operation)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.PictureFormat? pictureFormat = null;
            Office.Crop? crop = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null) return slideValidation;

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                var shapeValidation = ValidateShapeIndex(shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = shapes[shapeIndex];
                var typeValidation = ValidatePictureShape(shape, slideIndex, shapeIndex);
                if (typeValidation is not null) return typeValidation;

                pictureFormat = shape.PictureFormat;
                crop = pictureFormat.Crop;
                return operation(crop);
            }
            finally
            {
                ComUtilities.Release(ref crop);
                ComUtilities.Release(ref pictureFormat);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
        });
    }

    private static ImageOperationResult ReadCropFrame(int shapeIndex, Office.Crop crop) => new()
    {
        Success = true,
        ShapeIndex = shapeIndex,
        PictureWidth = crop.PictureWidth,
        PictureHeight = crop.PictureHeight,
        PictureOffsetX = crop.PictureOffsetX,
        PictureOffsetY = crop.PictureOffsetY,
        FrameLeft = crop.ShapeLeft,
        FrameTop = crop.ShapeTop,
        FrameWidth = crop.ShapeWidth,
        FrameHeight = crop.ShapeHeight
    };

    private static ImageOperationResult Failure(string message) => new()
    {
        Success = false,
        ErrorMessage = message
    };
}
