using Sbroenne.PowerPointMcp.ComInterop.Session;
using System.Linq;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <inheritdoc cref="ITextFrameCommands"/>
public sealed class TextFrameCommands : ITextFrameCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // PpBulletType (curated: only the members meaningful for a simple on/off toggle).
    private const int PpBulletNone = 0;
    private const int PpBulletUnnumbered = 1;

    // PpParagraphAlignment, verified against learn.microsoft.com/office/vba/api/powerpoint.ppparagraphalignment.
    // ppAlignmentMixed (-2) is intentionally excluded from the settable dictionary (it's a
    // read-only "mixed state across paragraphs" indicator, not a value you can set).
    private static readonly Dictionary<string, int> ParagraphAlignments = new()
    {
        ["ppAlignLeft"] = 1,
        ["ppAlignCenter"] = 2,
        ["ppAlignRight"] = 3,
        ["ppAlignJustify"] = 4,
        ["ppAlignDistribute"] = 5,
        ["ppAlignThaiDistribute"] = 6,
        ["ppAlignJustifyLow"] = 7,
    };

    private static readonly Dictionary<int, string> ParagraphAlignmentsByValue =
        ParagraphAlignments.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    /// <inheritdoc/>
    public TextFrameOperationResult SetText(IPresentationBatch batch, int slideIndex, int shapeIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateIndices(ctx, slideIndex, shapeIndex);
            if (validation is not null) return validation;

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            // Font.Bold is typed as Microsoft.Office.Core.MsoTriState (office.dll) — set via
            // dynamic with the raw int constant, same pattern as elsewhere in this project.
            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Bold = bold ? MsoTrue : MsoFalse;

            return new TextFrameOperationResult { Success = true, Bold = bold };
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Italic = italic ? MsoTrue : MsoFalse;

            return new TextFrameOperationResult { Success = true, Italic = italic };
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            bool italic = (int)shape.TextFrame.TextRange.Font.Italic == MsoTrue;

            return new TextFrameOperationResult { Success = true, Italic = italic };
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            shape.TextFrame.TextRange.Font.Underline = underline ? MsoTrue : MsoFalse;

            return new TextFrameOperationResult { Success = true, Underline = underline };
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            bool underline = (int)shape.TextFrame.TextRange.Font.Underline == MsoTrue;

            return new TextFrameOperationResult { Success = true, Underline = underline };
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            int alignmentValue = (int)shape.TextFrame.TextRange.ParagraphFormat.Alignment;

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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic bullet = shape.TextFrame.TextRange.ParagraphFormat.Bullet;

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

            dynamic shape = ctx.Presentation.Slides[slideIndex].Shapes[shapeIndex];
            dynamic bullet = shape.TextFrame.TextRange.ParagraphFormat.Bullet;
            int bulletType = (int)bullet.Type;
            bool enabled = bulletType != PpBulletNone;

            return new TextFrameOperationResult
            {
                Success = true,
                BulletEnabled = enabled,
                BulletCharacter = enabled ? char.ConvertFromUtf32((int)bullet.Character) : null
            };
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

        int shapeCount = ctx.Presentation.Slides[slideIndex].Shapes.Count;
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
