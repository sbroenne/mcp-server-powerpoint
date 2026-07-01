using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <inheritdoc cref="ITextFrameCommands"/>
public sealed class TextFrameCommands : ITextFrameCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

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
