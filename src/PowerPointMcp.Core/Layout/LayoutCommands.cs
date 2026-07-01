using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <inheritdoc cref="ILayoutCommands"/>
public sealed class LayoutCommands : ILayoutCommands
{
    /// <inheritdoc/>
    public LayoutOperationResult SetLayout(IPresentationBatch batch, int slideIndex, string layoutName)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(layoutName);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            if (!Enum.TryParse<PowerPoint.PpSlideLayout>(layoutName, ignoreCase: true, out var layout))
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{layoutName}' is not a recognized PpSlideLayout name (e.g. 'ppLayoutBlank', 'ppLayoutTitleOnly', 'ppLayoutText')."
                };
            }

            ctx.Presentation.Slides[slideIndex].Layout = layout;

            return new LayoutOperationResult { Success = true, LayoutName = layout.ToString() };
        });
    }

    /// <inheritdoc/>
    public LayoutOperationResult GetLayout(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            PowerPoint.PpSlideLayout layout = ctx.Presentation.Slides[slideIndex].Layout;

            return new LayoutOperationResult { Success = true, LayoutName = layout.ToString() };
        });
    }
}
