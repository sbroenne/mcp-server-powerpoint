using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <inheritdoc cref="ISlideCommands"/>
public sealed class SlideCommands : ISlideCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

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

    /// <inheritdoc/>
    public SlideOperationResult Duplicate(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int count = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            // Duplicate() returns a SlideRange containing the single new slide, inserted
            // immediately after the source slide.
            dynamic duplicateRange = ctx.Presentation.Slides[slideIndex].Duplicate();
            int newSlideIndex = (int)duplicateRange[1].SlideIndex;

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = newSlideIndex,
                SlideCount = ctx.Presentation.Slides.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult MoveTo(IPresentationBatch batch, int slideIndex, int toPosition)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int count = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            if (toPosition < 1 || toPosition > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Target position {toPosition} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            ctx.Presentation.Slides[slideIndex].MoveTo(toPosition);

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = toPosition,
                SlideCount = count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult SetBackgroundColor(IPresentationBatch batch, int slideIndex, byte red, byte green, byte blue)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int count = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            slide.FollowMasterBackground = MsoFalse;
            int rgb = red + (green << 8) + (blue << 16);
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = rgb;

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = slideIndex,
                ColorRgb = rgb,
                FollowsMasterBackground = false
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetBackgroundColor(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int count = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {count} slide(s) (valid range: 1-{count}).",
                    SlideCount = count
                };
            }

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            bool followsMaster = (int)slide.FollowMasterBackground == MsoTrue;
            int rgb = (int)slide.Background.Fill.ForeColor.RGB;

            return new SlideOperationResult
            {
                Success = true,
                SlideIndex = slideIndex,
                ColorRgb = rgb,
                FollowsMasterBackground = followsMaster
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult AddSection(IPresentationBatch batch, int sectionIndex, string? sectionName = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = (int)sectionProperties.Count;

            if (sectionIndex < 1 || sectionIndex > currentCount + 1)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Section index {sectionIndex} is out of range. The presentation has {currentCount} section(s) (valid range: 1-{currentCount + 1}).",
                    SectionCount = currentCount
                };
            }

            int newSectionIndex = sectionName is null
                ? (int)sectionProperties.AddSection(sectionIndex)
                : (int)sectionProperties.AddSection(sectionIndex, sectionName);

            return new SlideOperationResult
            {
                Success = true,
                SectionIndex = newSectionIndex,
                SectionCount = (int)sectionProperties.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult RenameSection(IPresentationBatch batch, int sectionIndex, string sectionName)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(sectionName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = (int)sectionProperties.Count;

            if (sectionIndex < 1 || sectionIndex > currentCount)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Section index {sectionIndex} is out of range. The presentation has {currentCount} section(s) (valid range: 1-{currentCount}).",
                    SectionCount = currentCount
                };
            }

            sectionProperties.Rename(sectionIndex, sectionName);

            return new SlideOperationResult
            {
                Success = true,
                SectionIndex = sectionIndex,
                SectionName = sectionName,
                SectionCount = currentCount
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult DeleteSection(IPresentationBatch batch, int sectionIndex, bool deleteSlides = false)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = (int)sectionProperties.Count;

            if (sectionIndex < 1 || sectionIndex > currentCount)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Section index {sectionIndex} is out of range. The presentation has {currentCount} section(s) (valid range: 1-{currentCount}).",
                    SectionCount = currentCount
                };
            }

            // The Delete method's second parameter is a plain VARIANT_BOOL, not an MsoTriState —
            // pass a real bool rather than the MsoTrue/MsoFalse Long constants used elsewhere.
            // Note: PowerPoint disallows deleting section 1 unless deleteSlides is true.
            sectionProperties.Delete(sectionIndex, deleteSlides);

            return new SlideOperationResult
            {
                Success = true,
                SectionIndex = sectionIndex,
                SectionCount = (int)sectionProperties.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetSectionCount(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sectionProperties = ctx.Presentation.SectionProperties;
            return new SlideOperationResult
            {
                Success = true,
                SectionCount = (int)sectionProperties.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetSectionName(IPresentationBatch batch, int sectionIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = (int)sectionProperties.Count;

            if (sectionIndex < 1 || sectionIndex > currentCount)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Section index {sectionIndex} is out of range. The presentation has {currentCount} section(s) (valid range: 1-{currentCount}).",
                    SectionCount = currentCount
                };
            }

            string name = (string)sectionProperties.Name(sectionIndex);

            return new SlideOperationResult
            {
                Success = true,
                SectionIndex = sectionIndex,
                SectionName = name,
                SectionCount = currentCount
            };
        });
    }
}
