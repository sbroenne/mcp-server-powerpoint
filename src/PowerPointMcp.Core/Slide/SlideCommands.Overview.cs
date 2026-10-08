extern alias OfficeInterop;

using System.Text;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <inheritdoc cref="ISlideCommands.Inspect"/>
public sealed partial class SlideCommands
{
    private const int DefaultOverviewSlideLimit = 20;
    private const int MaximumOverviewSlideLimit = 100;
    private const int MaximumOverviewShapesPerSlide = 200;
    private const int DefaultOverviewTextLimit = 500;
    private const int MaximumOverviewTextLimit = 2000;

    /// <inheritdoc/>
    public SlideOperationResult Inspect(
        IPresentationBatch batch,
        int maxSlides = DefaultOverviewSlideLimit,
        int maxTextCharsPerSlide = DefaultOverviewTextLimit)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (maxSlides is < 1 or > MaximumOverviewSlideLimit)
        {
            return new SlideOperationResult
            {
                ErrorMessage = $"maxSlides must be between 1 and {MaximumOverviewSlideLimit}."
            };
        }

        if (maxTextCharsPerSlide is < 0 or > MaximumOverviewTextLimit)
        {
            return new SlideOperationResult
            {
                ErrorMessage = $"maxTextCharsPerSlide must be between 0 and {MaximumOverviewTextLimit}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideCount = slides.Count;

                var overview = new List<SlideOverviewInfo>(Math.Min(slideCount, maxSlides));

                for (var slideIndex = 1; slideIndex <= Math.Min(slideCount, maxSlides); slideIndex++)
                {
                    ct.ThrowIfCancellationRequested();
                    PowerPoint.Slide? slide = null;
                    PowerPoint.CustomLayout? layout = null;
                    PowerPoint.Shapes? shapes = null;
                    try
                    {
                        slide = slides[slideIndex];
                        layout = slide.CustomLayout;
                        shapes = slide.Shapes;

                        string? preview = null;
                        var truncated = false;
                        if (maxTextCharsPerSlide > 0)
                        {
                            preview = ReadTextPreview(
                                shapes,
                                maxTextCharsPerSlide,
                                MaximumOverviewShapesPerSlide,
                                ct,
                                out truncated);
                        }

                        overview.Add(new SlideOverviewInfo
                        {
                            SlideIndex = slideIndex,
                            Name = slide.Name,
                            LayoutName = layout.Name,
                            ShapeCount = shapes.Count,
                            TextPreview = preview,
                            TextTruncated = truncated
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes!);
                        ComUtilities.Release(ref layout!);
                        ComUtilities.Release(ref slide!);
                    }
                }

                return new SlideOperationResult
                {
                    Success = true,
                    SlideCount = slideCount,
                    Slides = overview,
                    OmittedSlideCount = slideCount - overview.Count
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }

    private static string? ReadTextPreview(
        PowerPoint.Shapes shapes,
        int maximumCharacters,
        int maximumShapeCount,
        CancellationToken cancellationToken,
        out bool truncated)
    {
        var preview = new StringBuilder(Math.Min(maximumCharacters, DefaultOverviewTextLimit));
        truncated = false;
        var shapeCount = shapes.Count;

        for (var shapeIndex = 1; shapeIndex <= Math.Min(shapeCount, maximumShapeCount); shapeIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPoint.Shape? shape = null;
            PowerPoint.TextFrame? textFrame = null;
            PowerPoint.TextRange? textRange = null;
            PowerPoint.TextRange? previewRange = null;
            try
            {
                shape = shapes[shapeIndex];
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    continue;
                }

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                var remaining = maximumCharacters - preview.Length;
                var separatorLength = preview.Length == 0 ? 0 : 1;
                if (remaining <= separatorLength)
                {
                    truncated = true;
                    break;
                }

                var available = remaining - separatorLength;
                var textLength = textRange.Length;
                if (textLength == 0)
                {
                    continue;
                }

                var textLengthToRead = Math.Min(textLength, available + 1);
                previewRange = textRange.Characters(1, textLengthToRead);
                var shapeText = ((string)previewRange.Text).Replace("\r\n", "\n", StringComparison.Ordinal)
                    .Replace('\r', '\n')
                    .Trim();
                if (shapeText.Length == 0)
                {
                    if (textLength > textLengthToRead)
                    {
                        truncated = true;
                        break;
                    }
                    continue;
                }

                if (separatorLength > 0)
                {
                    preview.Append('\n');
                }

                if (shapeText.Length > available)
                {
                    preview.Append(shapeText, 0, available);
                    truncated = true;
                    break;
                }

                preview.Append(shapeText);
                if (textLength > textLengthToRead)
                {
                    truncated = true;
                    break;
                }
            }
            finally
            {
                ComUtilities.Release(ref previewRange!);
                ComUtilities.Release(ref textRange!);
                ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
            }
        }

        if (!truncated && shapeCount > maximumShapeCount)
        {
            truncated = true;
        }

        return preview.Length == 0 ? null : preview.ToString();
    }
}
