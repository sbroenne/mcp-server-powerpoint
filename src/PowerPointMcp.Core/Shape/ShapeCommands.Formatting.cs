using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    /// <inheritdoc/>
    public ShapeOperationResult CopyFormatting(
        IPresentationBatch batch,
        int slideIndex,
        int sourceShapeIndex,
        int targetShapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? sourceShape = null;
            PowerPoint.Shape? targetShape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null) return slideValidation;

                slide = slides[slideIndex];
                shapes = slide.Shapes;

                var sourceValidation = ValidateShapeIndex(shapes.Count, sourceShapeIndex);
                if (sourceValidation is not null) return sourceValidation;

                var targetValidation = ValidateShapeIndex(shapes.Count, targetShapeIndex);
                if (targetValidation is not null) return targetValidation;

                sourceShape = shapes[sourceShapeIndex];
                targetShape = shapes[targetShapeIndex];
                sourceShape.PickUp();
                targetShape.Apply();

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = targetShapeIndex
                };
            }
            finally
            {
                if (targetShape is not null) ComUtilities.Release(ref targetShape);
                if (sourceShape is not null) ComUtilities.Release(ref sourceShape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }
}
