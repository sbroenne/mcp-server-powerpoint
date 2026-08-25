using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Accessibility;

/// <inheritdoc cref="IAccessibilityCommands"/>
public sealed class AccessibilityCommands : IAccessibilityCommands
{
    private const int MsoTrue = -1;
    private const int MsoPicture = 13;
    private const int MsoLinkedPicture = 11;
    private const int MsoPlaceholder = 14;
    private const int MsoBringToFront = 0;

    /// <inheritdoc/>
    public AccessibilityOperationResult Audit(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var issues = new List<AccessibilityOperationResult.AccessibilityIssue>();
            PowerPoint.Slides? slides = null;
            try
            {
                slides = ctx.Presentation.Slides;
                int slideCount = slides.Count;
                for (int slideIndex = 1; slideIndex <= slideCount; slideIndex++)
                {
                    AuditSlide(slides, slideIndex, issues);
                }

                return new AccessibilityOperationResult
                {
                    Success = true,
                    Issues = issues
                };
            }
            finally
            {
                if (slides != null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public AccessibilityOperationResult GetReadingOrder(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            var validation = ValidateSlideIndex(slideCount, slideIndex);
            if (validation != null) return validation;

            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                shapes = slide.Shapes;
                var ownedShapes = new List<PowerPoint.Shape>(shapes.Count);
                try
                {
                    for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        ownedShapes.Add(shapes[shapeIndex]);
                    }

                    var stableShapes = ownedShapes.OrderBy(shape => shape.Id).ToArray();
                    return new AccessibilityOperationResult
                    {
                        Success = true,
                        ReadingOrder = stableShapes
                            .Select((shape, stableIndex) => new
                            {
                                ShapeIndex = stableIndex + 1,
                                shape.ZOrderPosition
                            })
                            .OrderBy(item => item.ZOrderPosition)
                            .Select(item => item.ShapeIndex)
                            .ToArray()
                    };
                }
                finally
                {
                    ReleaseShapes(ownedShapes);
                }
            }
            finally
            {
                if (shapes != null) ComUtilities.Release(ref shapes);
                if (slide != null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public AccessibilityOperationResult SetReadingOrder(
        IPresentationBatch batch,
        int slideIndex,
        IReadOnlyList<int> shapeIndexes)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(shapeIndexes);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            var slideValidation = ValidateSlideIndex(slideCount, slideIndex);
            if (slideValidation != null) return slideValidation;

            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            var orderedShapes = new List<PowerPoint.Shape>();
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                shapes = slide.Shapes;
                int shapeCount = shapes.Count;

                if (!IsCompletePermutation(shapeIndexes, shapeCount))
                {
                    return new AccessibilityOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Reading order must contain every 1-based shape index from 1 through {shapeCount} exactly once."
                    };
                }

                var shapesByStableIndex = new List<PowerPoint.Shape>(shapeCount);
                for (int shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex++)
                {
                    shapesByStableIndex.Add(shapes[shapeIndex]);
                }

                shapesByStableIndex.Sort((left, right) => left.Id.CompareTo(right.Id));
                foreach (int shapeIndex in shapeIndexes)
                {
                    orderedShapes.Add(shapesByStableIndex[shapeIndex - 1]);
                }

                foreach (PowerPoint.Shape shape in orderedShapes)
                {
                    // PIA gap: Shape.ZOrder takes Office.MsoZOrderCmd, unavailable without office.dll.
                    ((dynamic)shape).ZOrder(MsoBringToFront);
                }

                return new AccessibilityOperationResult
                {
                    Success = true,
                    ReadingOrder = shapeIndexes.ToArray()
                };
            }
            finally
            {
                ReleaseShapes(orderedShapes);
                if (shapes != null) ComUtilities.Release(ref shapes);
                if (slide != null) ComUtilities.Release(ref slide);
            }
        });
    }

    private static void AuditSlide(
        PowerPoint.Slides slides,
        int slideIndex,
        List<AccessibilityOperationResult.AccessibilityIssue> issues)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            slide = slides[slideIndex];
            shapes = slide.Shapes;
            var zOrderPositions = new HashSet<int>();

            for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[shapeIndex];
                    if (IsNonDecorativeVisualShape(shape) &&
                        string.IsNullOrWhiteSpace(shape.AlternativeText))
                    {
                        issues.Add(new AccessibilityOperationResult.AccessibilityIssue
                        {
                            Code = "missing-alt-text",
                            Message = $"Visual shape {shapeIndex} on slide {slideIndex} is missing alternative text.",
                            SlideIndex = slideIndex,
                            ShapeIndex = shapeIndex
                        });
                    }

                    if (IsEmptyTitlePlaceholder(shape))
                    {
                        issues.Add(new AccessibilityOperationResult.AccessibilityIssue
                        {
                            Code = "empty-title-placeholder",
                            Message = $"Title placeholder {shapeIndex} on slide {slideIndex} is empty.",
                            SlideIndex = slideIndex,
                            ShapeIndex = shapeIndex
                        });
                    }

                    int zOrderPosition = shape.ZOrderPosition;
                    if (zOrderPosition < 1 || zOrderPosition > shapes.Count)
                    {
                        issues.Add(new AccessibilityOperationResult.AccessibilityIssue
                        {
                            Code = "invalid-reading-order",
                            Message = $"Shape {shapeIndex} on slide {slideIndex} has invalid reading-order position {zOrderPosition}.",
                            SlideIndex = slideIndex,
                            ShapeIndex = shapeIndex
                        });
                    }
                    else if (!zOrderPositions.Add(zOrderPosition))
                    {
                        issues.Add(new AccessibilityOperationResult.AccessibilityIssue
                        {
                            Code = "duplicate-reading-order",
                            Message = $"Shape {shapeIndex} on slide {slideIndex} duplicates reading-order position {zOrderPosition}.",
                            SlideIndex = slideIndex,
                            ShapeIndex = shapeIndex
                        });
                    }
                }
                finally
                {
                    if (shape != null) ComUtilities.Release(ref shape);
                }
            }
        }
        finally
        {
            if (shapes != null) ComUtilities.Release(ref shapes);
            if (slide != null) ComUtilities.Release(ref slide);
        }
    }

    private static bool IsNonDecorativeVisualShape(PowerPoint.Shape shape)
    {
        // Type, HasChart, HasSmartArt, and Decorative are Office.Core-backed members, so these
        // reads are narrowly late-bound because office.dll is intentionally not referenced.
        dynamic dispatch = shape;
        int shapeType = (int)dispatch.Type;
        bool isVisual = shapeType is MsoPicture or MsoLinkedPicture ||
                        (int)dispatch.HasChart == MsoTrue ||
                        (int)dispatch.HasSmartArt == MsoTrue;
        return isVisual && (int)dispatch.Decorative != MsoTrue;
    }

    private static bool IsEmptyTitlePlaceholder(PowerPoint.Shape shape)
    {
        // PIA gap: Shape.Type is Office.MsoShapeType, unavailable without office.dll.
        if ((int)((dynamic)shape).Type != MsoPlaceholder)
        {
            return false;
        }

        PowerPoint.PlaceholderFormat? placeholderFormat = null;
        PowerPoint.TextFrame? textFrame = null;
        PowerPoint.TextRange? textRange = null;
        try
        {
            placeholderFormat = shape.PlaceholderFormat;
            PowerPoint.PpPlaceholderType placeholderType = placeholderFormat.Type;
            if (placeholderType is not PowerPoint.PpPlaceholderType.ppPlaceholderTitle and
                not PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle)
            {
                return false;
            }

            textFrame = shape.TextFrame;
            textRange = textFrame.TextRange;
            return string.IsNullOrWhiteSpace(textRange.Text);
        }
        finally
        {
            if (textRange != null) ComUtilities.Release(ref textRange);
            if (textFrame != null) ComUtilities.Release(ref textFrame);
            if (placeholderFormat != null) ComUtilities.Release(ref placeholderFormat);
        }
    }

    private static bool IsCompletePermutation(IReadOnlyList<int> shapeIndexes, int shapeCount)
    {
        if (shapeIndexes.Count != shapeCount)
        {
            return false;
        }

        var seen = new HashSet<int>();
        foreach (int shapeIndex in shapeIndexes)
        {
            if (shapeIndex < 1 || shapeIndex > shapeCount || !seen.Add(shapeIndex))
            {
                return false;
            }
        }

        return true;
    }

    private static void ReleaseShapes(List<PowerPoint.Shape> shapes)
    {
        foreach (PowerPoint.Shape shape in shapes)
        {
            PowerPoint.Shape? ownedShape = shape;
            ComUtilities.Release(ref ownedShape);
        }
    }

    private static AccessibilityOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new AccessibilityOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }

        return null;
    }
}
