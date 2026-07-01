using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <inheritdoc cref="ISlideCommands"/>
public sealed class SlideCommands : ISlideCommands
{
    /// <inheritdoc/>
    public SlideOperationResult AddBlank(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        // No try/catch suppression here (Rule 1b) — unexpected COM failures propagate to
        // batch.Execute()'s own exception handling.
        return batch.Execute((ctx, ct) =>
        {
            int newIndex = ctx.Presentation.Slides.Count + 1;
            ctx.Presentation.Slides.Add(newIndex, PowerPoint.PpSlideLayout.ppLayoutBlank);

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = newIndex,
                SlideCount = ctx.Presentation.Slides.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetCount(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) => new SlideOperationResult
        {
            Success = true,
            SlideCount = ctx.Presentation.Slides.Count
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult Delete(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int count = ctx.Presentation.Slides.Count;

            // Validate bounds explicitly and return a failure result rather than letting the
            // COM layer throw an "Integer out of range" COMException — this is input
            // validation producing a graceful error result, not suppression of an unexpected
            // failure (Rule 1b is about not catching unexpected exceptions, not about
            // skipping up-front validation).
            if (slideIndex < 1 || slideIndex > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            ctx.Presentation.Slides[slideIndex].Delete();

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = slideIndex,
                SlideCount = ctx.Presentation.Slides.Count
            };
        });
    }
}
