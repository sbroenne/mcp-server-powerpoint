using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    // Shape.Type and Fill.Type are Office enums unavailable without office.dll.
    private const int MsoShapePlaceholder = 14;
    private const int MsoFillPicture = 6;

    /// <inheritdoc/>
    public ShapeOperationResult ListPlaceholders(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null) return validation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var placeholders = new List<PlaceholderInfo>();
            for (int shapeIndex = 1; shapeIndex <= slide.Shapes.Count; shapeIndex++)
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = slide.Shapes[shapeIndex];
                    // PIA gap: Shape.Type is Office.MsoShapeType, unavailable without office.dll.
                    if ((int)((dynamic)shape).Type != MsoShapePlaceholder)
                    {
                        continue;
                    }

                    placeholders.Add(ReadPlaceholder(shape, shapeIndex));
                }
                finally
                {
                    if (shape is not null)
                    {
                        ComUtilities.Release(ref shape);
                    }
                }
            }

            return new ShapeOperationResult
            {
                Success = true,
                ShapeCount = slide.Shapes.Count,
                Placeholders = placeholders
            };
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetPlaceholderText(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape? shape = null;
            PowerPoint.TextFrame? textFrame = null;
            PowerPoint.TextRange? textRange = null;
            try
            {
                shape = slide.Shapes[shapeIndex];
                // PIA gap: Shape.Type is Office.MsoShapeType, unavailable without office.dll.
                if ((int)((dynamic)shape).Type != MsoShapePlaceholder)
                {
                    return new ShapeOperationResult
                    {
                        ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a native placeholder.",
                        ShapeIndex = shapeIndex
                    };
                }

                // PIA gap: Shape.HasTextFrame returns Office.MsoTriState, unavailable without office.dll.
                if ((int)((dynamic)shape).HasTextFrame != MsoTrue)
                {
                    return new ShapeOperationResult
                    {
                        ErrorMessage = $"Placeholder shape {shapeIndex} does not support text.",
                        ShapeIndex = shapeIndex,
                        PlaceholderType = shape.PlaceholderFormat.Type.ToString()
                    };
                }

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.Text = text;
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    PlaceholderType = shape.PlaceholderFormat.Type.ToString(),
                    PlaceholderText = text
                };
            }
            finally
            {
                if (textRange is not null)
                {
                    ComUtilities.Release(ref textRange);
                }
                if (textFrame is not null)
                {
                    ComUtilities.Release(ref textFrame);
                }
                if (shape is not null)
                {
                    ComUtilities.Release(ref shape);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetPlaceholderImage(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string imagePath)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(imagePath);

        string fullImagePath;
        try
        {
            fullImagePath = Path.GetFullPath(imagePath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return new ShapeOperationResult { ErrorMessage = $"Invalid image path: {ex.Message}" };
        }

        if (!File.Exists(fullImagePath))
        {
            return new ShapeOperationResult { ErrorMessage = $"Image file not found: {fullImagePath}." };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape? shape = null;
            dynamic? fill = null;
            try
            {
                shape = slide.Shapes[shapeIndex];
                // PIA gap: Shape.Type is Office.MsoShapeType, unavailable without office.dll.
                if ((int)((dynamic)shape).Type != MsoShapePlaceholder)
                {
                    return new ShapeOperationResult
                    {
                        ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a native placeholder.",
                        ShapeIndex = shapeIndex
                    };
                }

                PowerPoint.PpPlaceholderType placeholderType = shape.PlaceholderFormat.Type;
                if (!SupportsImage(placeholderType))
                {
                    return new ShapeOperationResult
                    {
                        ErrorMessage = $"Placeholder type {placeholderType} does not support image content.",
                        ShapeIndex = shapeIndex,
                        PlaceholderType = placeholderType.ToString()
                    };
                }

                // PIA gap: FillFormat is an Office type unavailable without office.dll.
                fill = ((dynamic)shape).Fill;
                fill.UserPicture(fullImagePath);
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    PlaceholderType = placeholderType.ToString(),
                    HasImage = true
                };
            }
            finally
            {
                if (fill is not null)
                {
                    ComUtilities.Release(ref fill);
                }
                if (shape is not null)
                {
                    ComUtilities.Release(ref shape);
                }
            }
        });
    }

    private static PlaceholderInfo ReadPlaceholder(PowerPoint.Shape shape, int shapeIndex)
    {
        PowerPoint.TextFrame? textFrame = null;
        PowerPoint.TextRange? textRange = null;
        dynamic? fill = null;
        try
        {
            bool hasText = false;
            string? text = null;
            // PIA gap: Shape.HasTextFrame returns Office.MsoTriState, unavailable without office.dll.
            if ((int)((dynamic)shape).HasTextFrame == MsoTrue)
            {
                textFrame = shape.TextFrame;
                // PIA gap: TextFrame.HasText returns Office.MsoTriState, unavailable without office.dll.
                if ((int)((dynamic)textFrame).HasText == MsoTrue)
                {
                    textRange = textFrame.TextRange;
                    text = textRange.Text;
                    hasText = !string.IsNullOrWhiteSpace(text);
                }
            }

            // PIA gap: FillFormat is an Office type unavailable without office.dll.
            fill = ((dynamic)shape).Fill;
            bool hasImage = (int)fill.Type == MsoFillPicture;
            return new PlaceholderInfo
            {
                ShapeIndex = shapeIndex,
                PlaceholderType = shape.PlaceholderFormat.Type.ToString(),
                Name = shape.Name,
                AltText = shape.AlternativeText,
                Left = shape.Left,
                Top = shape.Top,
                Width = shape.Width,
                Height = shape.Height,
                HasText = hasText,
                Text = text,
                HasImage = hasImage,
                ContentState = hasImage ? "image" : hasText ? "text" : "empty"
            };
        }
        finally
        {
            if (fill is not null)
            {
                ComUtilities.Release(ref fill);
            }
            if (textRange is not null)
            {
                ComUtilities.Release(ref textRange);
            }
            if (textFrame is not null)
            {
                ComUtilities.Release(ref textFrame);
            }
        }
    }

    private static bool SupportsImage(PowerPoint.PpPlaceholderType placeholderType)
        => placeholderType is
            PowerPoint.PpPlaceholderType.ppPlaceholderPicture or
            PowerPoint.PpPlaceholderType.ppPlaceholderBitmap or
            PowerPoint.PpPlaceholderType.ppPlaceholderObject or
            PowerPoint.PpPlaceholderType.ppPlaceholderMediaClip;
}
