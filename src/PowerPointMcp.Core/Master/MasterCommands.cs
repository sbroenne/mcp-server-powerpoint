using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Master;

/// <inheritdoc cref="IMasterCommands"/>
public sealed class MasterCommands : IMasterCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // MsoGradientStyle member name -> value, for SetGradientBackground/GetGradientBackground —
    // same table/verified behavior as SlideCommands.GradientStyles (FillFormat.TwoColorGradient
    // must be called BEFORE setting ForeColor/BackColor.RGB).
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
    public MasterOperationResult GetTitleFont(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var placeholder = FindPlaceholder(ctx, PowerPoint.PpPlaceholderType.ppPlaceholderTitle);
            if (placeholder is null)
            {
                return NotFound("title");
            }

            return ReadFont(placeholder);
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult SetTitleFont(
        IPresentationBatch batch,
        string? fontName = null,
        float? fontSize = null,
        bool? bold = null,
        byte? red = null,
        byte? green = null,
        byte? blue = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var placeholder = FindPlaceholder(ctx, PowerPoint.PpPlaceholderType.ppPlaceholderTitle);
            if (placeholder is null)
            {
                return NotFound("title");
            }

            return ApplyFont(placeholder, fontName, fontSize, bold, red, green, blue);
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult GetBodyFont(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var placeholder = FindPlaceholder(ctx, PowerPoint.PpPlaceholderType.ppPlaceholderBody);
            if (placeholder is null)
            {
                return NotFound("body");
            }

            return ReadFont(placeholder);
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult SetBodyFont(
        IPresentationBatch batch,
        string? fontName = null,
        float? fontSize = null,
        bool? bold = null,
        byte? red = null,
        byte? green = null,
        byte? blue = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var placeholder = FindPlaceholder(ctx, PowerPoint.PpPlaceholderType.ppPlaceholderBody);
            if (placeholder is null)
            {
                return NotFound("body");
            }

            return ApplyFont(placeholder, fontName, fontSize, bold, red, green, blue);
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult GetBackgroundColor(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic master = ctx.Presentation.SlideMaster;
            int rgb = (int)master.Background.Fill.ForeColor.RGB;

            return new MasterOperationResult { Success = true, ColorRgb = rgb };
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult SetBackgroundColor(IPresentationBatch batch, byte red, byte green, byte blue)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            // PowerPoint/Office RGB integers are packed as 0x00BBGGRR (matches the VBA RGB()
            // function), not the more common 0x00RRGGBB.
            int rgb = red + (green << 8) + (blue << 16);

            dynamic master = ctx.Presentation.SlideMaster;
            master.Background.Fill.Solid();
            master.Background.Fill.ForeColor.RGB = rgb;

            return new MasterOperationResult { Success = true, ColorRgb = rgb };
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult SetGradientBackground(
        IPresentationBatch batch,
        byte red1, byte green1, byte blue1,
        byte red2, byte green2, byte blue2,
        string gradientStyle = "msoGradientHorizontal",
        int gradientVariant = 1)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (!GradientStyles.TryGetValue(gradientStyle, out int styleValue))
        {
            return new MasterOperationResult
            {
                Success = false,
                ErrorMessage = $"Unrecognized gradientStyle '{gradientStyle}'. Valid values: {string.Join(", ", GradientStyles.Keys)}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int rgb1 = red1 + (green1 << 8) + (blue1 << 16);
            int rgb2 = red2 + (green2 << 8) + (blue2 << 16);

            dynamic master = ctx.Presentation.SlideMaster;
            // TwoColorGradient() must be called BEFORE setting ForeColor/BackColor — it resets
            // both colors to PowerPoint's defaults as a side effect (verified via diagnostic spike).
            master.Background.Fill.TwoColorGradient(styleValue, gradientVariant);
            master.Background.Fill.ForeColor.RGB = rgb1;
            master.Background.Fill.BackColor.RGB = rgb2;

            return new MasterOperationResult
            {
                Success = true,
                ColorRgb = rgb1,
                ColorRgb2 = rgb2,
                GradientStyleName = gradientStyle,
                GradientVariant = gradientVariant
            };
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult GetGradientBackground(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic master = ctx.Presentation.SlideMaster;
            int fillType = (int)master.Background.Fill.Type;
            const int MsoFillGradient = 3;
            if (fillType != MsoFillGradient)
            {
                return new MasterOperationResult
                {
                    Success = false,
                    ErrorMessage = $"The slide master's background fill is not a gradient (fill type = {fillType})."
                };
            }

            int rgb1 = (int)master.Background.Fill.ForeColor.RGB;
            int rgb2 = (int)master.Background.Fill.BackColor.RGB;
            int styleValue = (int)master.Background.Fill.GradientStyle;
            int variant = (int)master.Background.Fill.GradientVariant;
            string? styleName = GradientStylesByValue.GetValueOrDefault(styleValue);

            return new MasterOperationResult
            {
                Success = true,
                ColorRgb = rgb1,
                ColorRgb2 = rgb2,
                GradientStyleName = styleName,
                GradientVariant = variant
            };
        });
    }

    /// <summary>
    /// Finds the master placeholder shape of the given type by scanning the slide master's
    /// <c>Shapes</c> collection for a shape whose <c>PlaceholderFormat.Type</c> matches.
    /// Returns null (not an exception) if no such placeholder exists on this master — an expected
    /// condition for masters built from unusual/blank layouts, handled by callers as a validation
    /// failure (Rule 1b).
    /// </summary>
    private static dynamic? FindPlaceholder(PresentationContext ctx, PowerPoint.PpPlaceholderType type)
    {
        dynamic master = ctx.Presentation.SlideMaster;
        int shapeCount = (int)master.Shapes.Count;

        for (int i = 1; i <= shapeCount; i++)
        {
            dynamic shape = master.Shapes[i];
            bool hasPlaceholder = (int)shape.Type == 14 /* msoPlaceholder */;
            if (!hasPlaceholder)
            {
                continue;
            }

            var placeholderType = (PowerPoint.PpPlaceholderType)shape.PlaceholderFormat.Type;
            if (placeholderType == type)
            {
                return shape;
            }
        }

        return null;
    }

    private static MasterOperationResult ReadFont(dynamic placeholder)
    {
        string fontName = (string)placeholder.TextFrame.TextRange.Font.Name;
        float fontSize = (float)placeholder.TextFrame.TextRange.Font.Size;
        bool bold = (int)placeholder.TextFrame.TextRange.Font.Bold == MsoTrue;
        int colorRgb = (int)placeholder.TextFrame.TextRange.Font.Color.RGB;

        return new MasterOperationResult
        {
            Success = true,
            FontName = fontName,
            FontSize = fontSize,
            Bold = bold,
            ColorRgb = colorRgb
        };
    }

    private static MasterOperationResult ApplyFont(
        dynamic placeholder,
        string? fontName,
        float? fontSize,
        bool? bold,
        byte? red,
        byte? green,
        byte? blue)
    {
        if (fontName is not null)
        {
            placeholder.TextFrame.TextRange.Font.Name = fontName;
        }

        if (fontSize is not null)
        {
            placeholder.TextFrame.TextRange.Font.Size = fontSize.Value;
        }

        if (bold is not null)
        {
            placeholder.TextFrame.TextRange.Font.Bold = bold.Value ? MsoTrue : MsoFalse;
        }

        if (red is not null || green is not null || blue is not null)
        {
            // Missing channels default to 0 — callers are expected to pass all three together
            // when setting color (mirrors TextFrameCommands.SetFontColor's all-or-nothing shape).
            int rgb = (red ?? 0) + ((green ?? 0) << 8) + ((blue ?? 0) << 16);
            placeholder.TextFrame.TextRange.Font.Color.RGB = rgb;
        }

        // Re-read from the placeholder so the result reflects the values actually applied
        // (including any font properties left unchanged by this call).
        return ReadFont(placeholder);
    }

    private static MasterOperationResult NotFound(string placeholderName)
        => new()
        {
            Success = false,
            ErrorMessage = $"The slide master does not have a '{placeholderName}' placeholder."
        };
}
