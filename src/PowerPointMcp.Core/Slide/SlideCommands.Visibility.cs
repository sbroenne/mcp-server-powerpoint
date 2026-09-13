extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

public sealed partial class SlideCommands
{
    /// <inheritdoc/>
    public SlideOperationResult SetHidden(IPresentationBatch batch, int slideIndex, bool hidden)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.SlideShowTransition? transition = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var validation = ValidateSlideIndex(slides.Count, slideIndex);
                if (validation is not null) return validation;

                slide = slides[slideIndex];
                transition = slide.SlideShowTransition;
                transition.Hidden = hidden ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    SlideCount = slides.Count,
                    Hidden = transition.Hidden == Office.MsoTriState.msoTrue
                };
            }
            finally
            {
                ComUtilities.Release(ref transition!);
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult SetDisplayMasterShapes(IPresentationBatch batch, int slideIndex, bool display)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var validation = ValidateSlideIndex(slides.Count, slideIndex);
                if (validation is not null) return validation;

                slide = slides[slideIndex];
                slide.DisplayMasterShapes = display ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    SlideCount = slides.Count,
                    DisplaysMasterShapes = slide.DisplayMasterShapes == Office.MsoTriState.msoTrue
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }
}