using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <inheritdoc cref="ITextFrameCommands"/>
public sealed class TextFrameCommands : ITextFrameCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // PpBulletType (curated: only the members meaningful for a simple on/off toggle).
    private const PowerPoint.PpBulletType PpBulletNone = PowerPoint.PpBulletType.ppBulletNone;
    private const PowerPoint.PpBulletType PpBulletUnnumbered = PowerPoint.PpBulletType.ppBulletUnnumbered;

    // PpParagraphAlignment, verified against learn.microsoft.com/office/vba/api/powerpoint.ppparagraphalignment.
    // ppAlignmentMixed (-2) is intentionally excluded from the settable dictionary (it's a
    // read-only "mixed state across paragraphs" indicator, not a value you can set).
    private static readonly Dictionary<string, PowerPoint.PpParagraphAlignment> ParagraphAlignments = new()
    {
        ["ppAlignLeft"] = PowerPoint.PpParagraphAlignment.ppAlignLeft,
        ["ppAlignCenter"] = PowerPoint.PpParagraphAlignment.ppAlignCenter,
        ["ppAlignRight"] = PowerPoint.PpParagraphAlignment.ppAlignRight,
        ["ppAlignJustify"] = PowerPoint.PpParagraphAlignment.ppAlignJustify,
        ["ppAlignDistribute"] = PowerPoint.PpParagraphAlignment.ppAlignDistribute,
        ["ppAlignThaiDistribute"] = PowerPoint.PpParagraphAlignment.ppAlignThaiDistribute,
        ["ppAlignJustifyLow"] = PowerPoint.PpParagraphAlignment.ppAlignJustifyLow,
    };

    private static readonly Dictionary<PowerPoint.PpParagraphAlignment, string> ParagraphAlignmentsByValue =
        ParagraphAlignments.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    // PpAutoSize, verified against learn.microsoft.com/office/vba/api/powerpoint.ppautosize.
    // ppAutoSizeMixed (-2) is intentionally excluded from the settable dictionary (it's a
    // read-only "mixed state across shapes" indicator, not a value you can set).
    private static readonly Dictionary<string, PowerPoint.PpAutoSize> AutoSizeModes = new()
    {
        ["ppAutoSizeNone"] = PowerPoint.PpAutoSize.ppAutoSizeNone,
        ["ppAutoSizeShapeToFitText"] = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText,
        ["ppAutoSizeTextToFitShape"] = (PowerPoint.PpAutoSize)2,
    };

    private static readonly Dictionary<PowerPoint.PpAutoSize, string> AutoSizeModesByValue =
        AutoSizeModes.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    /// <inheritdoc/>
    public TextFrameOperationResult SetText(IPresentationBatch batch, int slideIndex, int shapeIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Text = text;

            return new TextFrameOperationResult { Success = true, Text = text };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetText(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            string text = (string)shape.TextFrame.TextRange.Text;

            return new TextFrameOperationResult { Success = true, Text = text };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex, float fontSize)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Size = fontSize;

            return new TextFrameOperationResult { Success = true, FontSize = fontSize };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetBold(IPresentationBatch batch, int slideIndex, int shapeIndex, bool bold)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic? font = null;
            try
            {
                // Font tri-state properties are Microsoft.Office.Core-typed; keep the property write late-bound.
                font = shape.TextFrame.TextRange.Font;
                font.Bold = bold ? MsoTrue : MsoFalse;

                return new TextFrameOperationResult { Success = true, Bold = bold };
            }
            finally
            {
                if (font != null)
                {
                    ComUtilities.Release(ref font!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            // PowerPoint/Office RGB integers are packed as 0x00BBGGRR (matches the VBA RGB()
            // function), not the more common 0x00RRGGBB.
            int rgb = red + (green << 8) + (blue << 16);

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Color.RGB = rgb;

            return new TextFrameOperationResult { Success = true, ColorRgb = rgb };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex, bool italic)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic? font = null;
            try
            {
                font = shape.TextFrame.TextRange.Font;
                font.Italic = italic ? MsoTrue : MsoFalse;

                return new TextFrameOperationResult { Success = true, Italic = italic };
            }
            finally
            {
                if (font != null)
                {
                    ComUtilities.Release(ref font!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic? font = null;
            try
            {
                font = shape.TextFrame.TextRange.Font;
                bool italic = (int)font.Italic == MsoTrue;

                return new TextFrameOperationResult { Success = true, Italic = italic };
            }
            finally
            {
                if (font != null)
                {
                    ComUtilities.Release(ref font!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex, bool underline)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic? font = null;
            try
            {
                font = shape.TextFrame.TextRange.Font;
                font.Underline = underline ? MsoTrue : MsoFalse;

                return new TextFrameOperationResult { Success = true, Underline = underline };
            }
            finally
            {
                if (font != null)
                {
                    ComUtilities.Release(ref font!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic? font = null;
            try
            {
                font = shape.TextFrame.TextRange.Font;
                bool underline = (int)font.Underline == MsoTrue;

                return new TextFrameOperationResult { Success = true, Underline = underline };
            }
            finally
            {
                if (font != null)
                {
                    ComUtilities.Release(ref font!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex, string fontName)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(fontName);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Name = fontName;

            return new TextFrameOperationResult { Success = true, FontName = fontName };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            string fontName = (string)shape.TextFrame.TextRange.Font.Name;

            return new TextFrameOperationResult { Success = true, FontName = fontName };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex, string alignment)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(alignment);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            if (!ParagraphAlignments.TryGetValue(alignment, out var alignmentValue))
            {
                return new TextFrameOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{alignment}' is not a recognized PpParagraphAlignment name (must be 'ppAlignLeft', 'ppAlignCenter', 'ppAlignRight', 'ppAlignJustify', or 'ppAlignDistribute')."
                };
            }

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = alignmentValue;

            return new TextFrameOperationResult { Success = true, Alignment = alignment };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            PowerPoint.PpParagraphAlignment alignmentValue = shape.TextFrame.TextRange.ParagraphFormat.Alignment;

            if (!ParagraphAlignmentsByValue.TryGetValue(alignmentValue, out var alignmentName))
            {
                return new TextFrameOperationResult
                {
                    Success = false,
                    ErrorMessage = $"The paragraph alignment value {alignmentValue} is not one of the recognized PpParagraphAlignment names (it may be mixed across multiple paragraphs)."
                };
            }

            return new TextFrameOperationResult { Success = true, Alignment = alignmentName };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex, bool enabled, string? character = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            if (character is { Length: > 1 })
            {
                return new TextFrameOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{character}' is not a single character. The bullet glyph must be exactly one character."
                };
            }

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            // Reason: ParagraphFormat.Bullet's embedded signature references Office.MsoTriState from
            // office.dll, which this project deliberately does not load. Keep only this boundary
            // late-bound; the containing PowerPoint shape and PpBulletType values remain typed.
            dynamic? bullet = null;
            try
            {
                // Reason: ParagraphFormat.Bullet's embedded signature references
                // Office.MsoTriState from office.dll, so it is read via dynamic late binding.
                bullet = ((dynamic)shape.TextFrame.TextRange.ParagraphFormat).Bullet;

                bullet.Type = enabled ? PpBulletUnnumbered : PpBulletNone;
                if (enabled && character is not null)
                {
                    bullet.Character = (int)character[0];
                }

                return new TextFrameOperationResult
                {
                    Success = true,
                    BulletEnabled = enabled,
                    BulletCharacter = enabled ? (character ?? char.ConvertFromUtf32((int)bullet.Character)) : null
                };
            }
            finally
            {
                if (bullet != null)
                {
                    ComUtilities.Release(ref bullet!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            // Reason: See SetBullet — the Bullet getter depends on Office.MsoTriState from office.dll.
            dynamic? bullet = null;
            try
            {
                // Reason: See SetBullet — the Bullet getter depends on Office.MsoTriState from office.dll.
                bullet = ((dynamic)shape.TextFrame.TextRange.ParagraphFormat).Bullet;
                bool enabled = (int)bullet.Type != (int)PpBulletNone;

                return new TextFrameOperationResult
                {
                    Success = true,
                    BulletEnabled = enabled,
                    BulletCharacter = enabled ? char.ConvertFromUtf32((int)bullet.Character) : null
                };
            }
            finally
            {
                if (bullet != null)
                {
                    ComUtilities.Release(ref bullet!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex, string autoSize)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(autoSize);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            if (!AutoSizeModes.TryGetValue(autoSize, out var autoSizeValue))
            {
                return new TextFrameOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{autoSize}' is not a recognized PpAutoSize name (must be 'ppAutoSizeNone', 'ppAutoSizeShapeToFitText', or 'ppAutoSizeTextToFitShape')."
                };
            }

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.AutoSize = autoSizeValue;

            return new TextFrameOperationResult { Success = true, AutoSize = autoSize };
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult GetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            PowerPoint.Shape shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            PowerPoint.PpAutoSize autoSizeValue = shape.TextFrame.AutoSize;

            if (!AutoSizeModesByValue.TryGetValue(autoSizeValue, out var autoSizeName))
            {
                return new TextFrameOperationResult
                {
                    Success = false,
                    ErrorMessage = $"The auto-size value {autoSizeValue} is not one of the recognized PpAutoSize names (it may be mixed across multiple shapes)."
                };
            }

            return new TextFrameOperationResult { Success = true, AutoSize = autoSizeName };
        });
    }

    private static TextFrameOperationResult? ValidateIndices(PresentationContext ctx, int slideIndex, int shapeIndex)
    {
        int slideCount = ctx.Presentation.Slides.Count;
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new TextFrameOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }

        PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
        int shapeCount = slide.Shapes.Count;
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new TextFrameOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }

        return null;
    }
}
