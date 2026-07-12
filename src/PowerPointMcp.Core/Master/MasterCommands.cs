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
            PowerPoint.Master master = GetSlideMaster(ctx);
            int rgb = master.Background.Fill.ForeColor.RGB;

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

            PowerPoint.Master master = GetSlideMaster(ctx);
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

            PowerPoint.Master master = GetSlideMaster(ctx);
            // TwoColorGradient() must be called BEFORE setting ForeColor/BackColor — it resets
            // both colors to PowerPoint's defaults as a side effect (verified via diagnostic spike).
            dynamic fill = master.Background.Fill;
            fill.TwoColorGradient(styleValue, gradientVariant);
            fill.ForeColor.RGB = rgb1;
            fill.BackColor.RGB = rgb2;

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
            PowerPoint.Master master = GetSlideMaster(ctx);
            dynamic fill = master.Background.Fill;
            int fillType = (int)fill.Type;
            const int MsoFillGradient = 3;
            if (fillType != MsoFillGradient)
            {
                return new MasterOperationResult
                {
                    Success = false,
                    ErrorMessage = $"The slide master's background fill is not a gradient (fill type = {fillType})."
                };
            }

            int rgb1 = (int)fill.ForeColor.RGB;
            int rgb2 = (int)fill.BackColor.RGB;
            int styleValue = (int)fill.GradientStyle;
            int variant = (int)fill.GradientVariant;
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
    private static PowerPoint.Shape? FindPlaceholder(PresentationContext ctx, PowerPoint.PpPlaceholderType type)
    {
        PowerPoint.Master master = GetSlideMaster(ctx);
        int shapeCount = master.Shapes.Count;

        for (int i = 1; i <= shapeCount; i++)
        {
            PowerPoint.Shape shape = master.Shapes[i];
            bool hasPlaceholder = (int)((dynamic)shape).Type == 14 /* msoPlaceholder */;
            if (!hasPlaceholder)
            {
                continue;
            }

            PowerPoint.PpPlaceholderType placeholderType = shape.PlaceholderFormat.Type;
            if (placeholderType == type)
            {
                return shape;
            }
        }

        return null;
    }

    private static PowerPoint.Master GetSlideMaster(PresentationContext ctx)
    {
        // The embedded NoPIA getter for Presentation.SlideMaster hangs indefinitely in live
        // PowerPoint. Dispatch only that getter through IDispatch, then return to typed PIA access.
        dynamic presentation = ctx.Presentation;
        return (PowerPoint.Master)presentation.SlideMaster;
    }

    private static MasterOperationResult ReadFont(PowerPoint.Shape placeholder)
    {
        dynamic font = placeholder.TextFrame.TextRange.Font;
        string fontName = (string)font.Name;
        float fontSize = (float)font.Size;
        bool bold = (int)font.Bold == MsoTrue;
        int colorRgb = (int)font.Color.RGB;

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
        PowerPoint.Shape placeholder,
        string? fontName,
        float? fontSize,
        bool? bold,
        byte? red,
        byte? green,
        byte? blue)
    {
        dynamic font = placeholder.TextFrame.TextRange.Font;

        if (fontName is not null)
        {
            font.Name = fontName;
        }

        if (fontSize is not null)
        {
            font.Size = fontSize.Value;
        }

        if (bold is not null)
        {
            font.Bold = bold.Value ? MsoTrue : MsoFalse;
        }

        if (red is not null || green is not null || blue is not null)
        {
            // Missing channels default to 0 — callers are expected to pass all three together
            // when setting color (mirrors TextFrameCommands.SetFontColor's all-or-nothing shape).
            int rgb = (red ?? 0) + ((green ?? 0) << 8) + ((blue ?? 0) << 16);
            font.Color.RGB = rgb;
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
