using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <inheritdoc cref="ISlideCommands"/>
public sealed class SlideCommands : ISlideCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // MsoGradientStyle member name -> value, for SetGradientBackground/GetGradientBackground
    // (learn.microsoft.com/office/vba/api/office.msogradientstyle) — verified live via a
    // temporary diagnostic spike (since removed): FillFormat.TwoColorGradient(style, variant)
    // must be called BEFORE setting ForeColor/BackColor.RGB, since TwoColorGradient() itself
    // resets both colors to PowerPoint's defaults (white/theme accent) as a side effect.
    private static readonly Dictionary<string, int> GradientStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoGradientHorizontal"] = 1,
        ["msoGradientVertical"] = 2,
        ["msoGradientDiagonalUp"] = 3,
        ["msoGradientDiagonalDown"] = 4,
        ["msoGradientFromCorner"] = 5,
        ["msoGradientFromTitle"] = 6,
        ["msoGradientFromCenter"] = 7,
    };

    private static readonly Dictionary<int, string> GradientStylesByValue =
        GradientStyles.ToDictionary(kv => kv.Value, kv => kv.Key);

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
            PowerPoint.SlideRange? duplicateRange = null;
            int newSlideIndex;
            try
            {
                duplicateRange = ctx.Presentation.Slides[slideIndex].Duplicate();
                newSlideIndex = duplicateRange[1].SlideIndex;
            }
            finally
            {
                if (duplicateRange != null)
                {
                    ComUtilities.Release(ref duplicateRange!);
                }
            }

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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic? dynSlide = null;
            PowerPoint.ShapeRange? background = null;
            dynamic? fill = null;
            try
            {
                dynSlide = slide;
                dynSlide.FollowMasterBackground = MsoFalse;
                int rgb = red + (green << 8) + (blue << 16);
                background = slide.Background;
                fill = background.Fill;
                fill.Solid();
                fill.ForeColor.RGB = rgb;

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    ColorRgb = rgb,
                    FollowsMasterBackground = false
                };
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }
                if (background != null)
                {
                    ComUtilities.Release(ref background!);
                }
                if (dynSlide != null)
                {
                    ComUtilities.Release(ref dynSlide!);
                }
            }
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic? dynSlide = null;
            PowerPoint.ShapeRange? background = null;
            dynamic? fill = null;
            try
            {
                dynSlide = slide;
                bool followsMaster = (int)dynSlide.FollowMasterBackground == MsoTrue;
                background = slide.Background;
                fill = background.Fill;
                int rgb = (int)fill.ForeColor.RGB;

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    ColorRgb = rgb,
                    FollowsMasterBackground = followsMaster
                };
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }
                if (background != null)
                {
                    ComUtilities.Release(ref background!);
                }
                if (dynSlide != null)
                {
                    ComUtilities.Release(ref dynSlide!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult SetGradientBackground(
        IPresentationBatch batch,
        int slideIndex,
        byte red1, byte green1, byte blue1,
        byte red2, byte green2, byte blue2,
        string gradientStyle = "msoGradientHorizontal",
        int gradientVariant = 1)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (!GradientStyles.TryGetValue(gradientStyle, out int styleValue))
        {
            return new SlideOperationResult
            {
                Success = false,
                ErrorMessage = $"Unrecognized gradientStyle '{gradientStyle}'. Valid values: {string.Join(", ", GradientStyles.Keys)}."
            };
        }

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

            int rgb1 = red1 + (green1 << 8) + (blue1 << 16);
            int rgb2 = red2 + (green2 << 8) + (blue2 << 16);

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic? dynSlide = null;
            PowerPoint.ShapeRange? background = null;
            dynamic? fill = null;
            try
            {
                dynSlide = slide;
                dynSlide.FollowMasterBackground = MsoFalse;
                // TwoColorGradient() must be called BEFORE setting ForeColor/BackColor — it resets
                // both colors to PowerPoint's defaults as a side effect (verified via diagnostic spike).
                background = slide.Background;
                fill = background.Fill;
                fill.TwoColorGradient(styleValue, gradientVariant);
                fill.ForeColor.RGB = rgb1;
                fill.BackColor.RGB = rgb2;

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    ColorRgb = rgb1,
                    ColorRgb2 = rgb2,
                    GradientStyleName = gradientStyle,
                    GradientVariant = gradientVariant,
                    FollowsMasterBackground = false
                };
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }

                if (background != null)
                {
                    ComUtilities.Release(ref background!);
                }

                if (dynSlide != null)
                {
                    ComUtilities.Release(ref dynSlide!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetGradientBackground(IPresentationBatch batch, int slideIndex)
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            PowerPoint.ShapeRange? background = null;
            dynamic? fill = null;
            try
            {
                background = slide.Background;
                fill = background.Fill;
                int fillType = (int)fill.Type;
                const int MsoFillGradient = 3;
                if (fillType != MsoFillGradient)
                {
                    return new SlideOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Slide {slideIndex}'s background fill is not a gradient (fill type = {fillType}).",
                        SlideIndex = slideIndex
                    };
                }

                int rgb1 = (int)fill.ForeColor.RGB;
                int rgb2 = (int)fill.BackColor.RGB;
                int styleValue = (int)fill.GradientStyle;
                int variant = (int)fill.GradientVariant;
                string? styleName = GradientStylesByValue.GetValueOrDefault(styleValue);

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    ColorRgb = rgb1,
                    ColorRgb2 = rgb2,
                    GradientStyleName = styleName,
                    GradientVariant = variant,
                    FollowsMasterBackground = false
                };
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }
                if (background != null)
                {
                    ComUtilities.Release(ref background!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult AddSection(IPresentationBatch batch, int sectionIndex, string? sectionName = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SectionProperties sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = sectionProperties.Count;

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
                ? sectionProperties.AddSection(sectionIndex)
                : sectionProperties.AddSection(sectionIndex, sectionName);

            return new SlideOperationResult
            {
                Success = true,
                SectionIndex = newSectionIndex,
                SectionCount = sectionProperties.Count
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
            PowerPoint.SectionProperties sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = sectionProperties.Count;

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
            PowerPoint.SectionProperties sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = sectionProperties.Count;

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
                SectionCount = sectionProperties.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetSectionCount(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SectionProperties sectionProperties = ctx.Presentation.SectionProperties;
            return new SlideOperationResult
            {
                Success = true,
                SectionCount = sectionProperties.Count
            };
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult GetSectionName(IPresentationBatch batch, int sectionIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SectionProperties sectionProperties = ctx.Presentation.SectionProperties;
            int currentCount = sectionProperties.Count;

            if (sectionIndex < 1 || sectionIndex > currentCount)
            {
                return new SlideOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Section index {sectionIndex} is out of range. The presentation has {currentCount} section(s) (valid range: 1-{currentCount}).",
                    SectionCount = currentCount
                };
            }

            string name = sectionProperties.Name(sectionIndex);

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
