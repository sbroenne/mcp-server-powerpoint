using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Tags;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

public sealed partial class SlideCommands
{
    /// <inheritdoc/>
    public SlideOperationResult SetTag(IPresentationBatch batch, int slideIndex, string tagName, string tagValue)
    {
        ArgumentNullException.ThrowIfNull(tagValue);
        return ExecuteTagOperation(
            batch,
            slideIndex,
            (ctx, slide, acquireTags) =>
                TagCollectionHelper.Set(ctx, slide, acquireTags, tagName, tagValue));
    }

    /// <inheritdoc/>
    public SlideOperationResult GetTag(IPresentationBatch batch, int slideIndex, string tagName) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            (ctx, slide, acquireTags) =>
                TagCollectionHelper.Get(ctx, slide, acquireTags, tagName));

    /// <inheritdoc/>
    public SlideOperationResult ListTags(IPresentationBatch batch, int slideIndex) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            (ctx, slide, acquireTags) => TagCollectionHelper.List(ctx, slide, acquireTags));

    /// <inheritdoc/>
    public SlideOperationResult DeleteTag(IPresentationBatch batch, int slideIndex, string tagName) =>
        ExecuteTagOperation(
            batch,
            slideIndex,
            (ctx, slide, acquireTags) =>
                TagCollectionHelper.Delete(ctx, slide, acquireTags, tagName));

    private static SlideOperationResult ExecuteTagOperation(
        IPresentationBatch batch,
        int slideIndex,
        Func<PresentationContext, PowerPoint.Slide, Func<PowerPoint.Tags>, TagCollectionResult> operation)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null)
            {
                return validation;
            }

            PowerPoint.Slide? slide = ctx.Presentation.Slides[slideIndex];
            bool ownerRetained = ctx.RetainOwnedComOwner(slide);
            try
            {
                TagCollectionResult result = operation(ctx, slide, () => slide.Tags);
                return new SlideOperationResult
                {
                    Success = result.Success,
                    ErrorMessage = result.ErrorMessage,
                    SlideIndex = slideIndex,
                    TagName = result.Name,
                    TagValue = result.Value,
                    TagIndex = result.Index,
                    TagCount = result.Count,
                    Tags = result.Tags
                };
            }
            finally
            {
                if (!ownerRetained)
                {
                    ComUtilities.Release(ref slide);
                }
            }
        });
    }
}
