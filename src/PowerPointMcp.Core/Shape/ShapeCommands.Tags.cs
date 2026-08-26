using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Tags;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    /// <inheritdoc/>
    public ShapeOperationResult SetTag(IPresentationBatch batch, int slideIndex, int shapeIndex, string tagName, string tagValue)
    {
        ArgumentNullException.ThrowIfNull(tagValue);
        return ExecuteTagOperation(
            batch,
            slideIndex,
            shapeIndex,
            (ctx, shape, acquireTags) =>
                TagCollectionHelper.Set(ctx, shape, acquireTags, tagName, tagValue));
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetTag(IPresentationBatch batch, int slideIndex, int shapeIndex, string tagName) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            shapeIndex,
            (ctx, shape, acquireTags) =>
                TagCollectionHelper.Get(ctx, shape, acquireTags, tagName));

    /// <inheritdoc/>
    public ShapeOperationResult ListTags(IPresentationBatch batch, int slideIndex, int shapeIndex) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            shapeIndex,
            (ctx, shape, acquireTags) => TagCollectionHelper.List(ctx, shape, acquireTags));

    /// <inheritdoc/>
    public ShapeOperationResult DeleteTag(IPresentationBatch batch, int slideIndex, int shapeIndex, string tagName) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            shapeIndex,
            (ctx, shape, acquireTags) =>
                TagCollectionHelper.Delete(ctx, shape, acquireTags, tagName));

    private static ShapeOperationResult ExecuteTagOperation(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        Func<PresentationContext, PowerPoint.Shape, Func<PowerPoint.Tags>, TagCollectionResult> operation)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null)
            {
                return slideValidation;
            }

            PowerPoint.Slide? slide = ctx.Presentation.Slides[slideIndex];
            bool slideRetained = ctx.RetainOwnedComOwner(slide);
            PowerPoint.Shape? shape = null;
            bool shapeRetained = false;
            try
            {
                var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null)
                {
                    return shapeValidation;
                }

                shape = slide.Shapes[shapeIndex];
                shapeRetained = ctx.RetainOwnedComOwner(shape);
                TagCollectionResult result = operation(ctx, shape, () => shape.Tags);
                return new ShapeOperationResult
                {
                    Success = result.Success,
                    ErrorMessage = result.ErrorMessage,
                    ShapeIndex = shapeIndex,
                    TagName = result.Name,
                    TagValue = result.Value,
                    TagIndex = result.Index,
                    TagCount = result.Count,
                    Tags = result.Tags
                };
            }
            finally
            {
                if (shape is not null && !shapeRetained)
                {
                    ComUtilities.Release(ref shape);
                }
                if (!slideRetained)
                {
                    ComUtilities.Release(ref slide);
                }
            }
        });
    }
}
