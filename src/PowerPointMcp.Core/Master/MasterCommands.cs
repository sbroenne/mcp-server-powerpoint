using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
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
            dynamic? fill = null;
            try
            {
                fill = master.Background.Fill;
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
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult GetGradientBackground(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Master master = GetSlideMaster(ctx);
            dynamic? fill = null;
            try
            {
                fill = master.Background.Fill;
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
            }
            finally
            {
                if (fill != null)
                {
                    ComUtilities.Release(ref fill!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult ListMasters(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Designs designs = ctx.Presentation.Designs;
            int designCount = designs.Count;
            var masters = new List<MasterOperationResult.MasterInventoryEntry>(designCount);

            for (int i = 1; i <= designCount; i++)
            {
                PowerPoint.Design design = designs[i];
                PowerPoint.Master slideMaster = design.SlideMaster;
                PowerPoint.CustomLayouts customLayouts = slideMaster.CustomLayouts;
                int layoutCount = customLayouts.Count;
                var layouts = new List<MasterOperationResult.LayoutInventoryEntry>(layoutCount);

                for (int j = 1; j <= layoutCount; j++)
                {
                    PowerPoint.CustomLayout layout = customLayouts[j];
                    layouts.Add(new MasterOperationResult.LayoutInventoryEntry
                    {
                        LayoutIndex = j,
                        LayoutName = GetLayoutName(layout),
                        IsUsed = IsLayoutUsedByAnySlide(ctx, design, layout)
                    });
                }

                masters.Add(new MasterOperationResult.MasterInventoryEntry
                {
                    MasterIndex = i,
                    MasterName = GetMasterName(design, slideMaster),
                    Layouts = layouts
                });
            }

            return new MasterOperationResult { Success = true, Masters = masters };
        });
    }

    /// <inheritdoc/>
    public MasterOperationResult DeleteMaster(IPresentationBatch batch, int masterIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (masterIndex < 1)
        {
            return new MasterOperationResult
            {
                Success = false,
                ErrorMessage = "Master index must be 1 or greater."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Designs designs = ctx.Presentation.Designs;
            int designCount = designs.Count;
            if (masterIndex > designCount)
            {
                return new MasterOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Master index {masterIndex} is out of range. The presentation has {designCount} master(s) (valid range: 1-{designCount})."
                };
            }

            PowerPoint.Design design = designs[masterIndex];
            if (IsMasterUsedByAnySlide(ctx, design))
            {
                return new MasterOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Master '{GetMasterName(design, design.SlideMaster)}' is still used by one or more slides."
                };
            }

            InvokeComMethod(design, "Delete");
            return new MasterOperationResult { Success = true };
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
            // Reason: Shape.Type is Microsoft.Office.Core.MsoShapeType (Office.Core — not embedded),
            // so it is read via dynamic late binding here.
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
        // NOTE: `presentation` is a dynamic alias of the shared, session-owned ctx.Presentation
        // object (not a new/owned COM reference) — it must NOT be released here, since doing so
        // decrements the reference count of the Presentation RCW shared by the whole batch/session.
        dynamic presentation = ctx.Presentation;
        return (PowerPoint.Master)presentation.SlideMaster;
    }

    private static bool IsLayoutUsedByAnySlide(PresentationContext ctx, PowerPoint.Design design, PowerPoint.CustomLayout layout)
    {
        string? targetDesignName = NormalizeName(design.Name);
        string? targetLayoutName = NormalizeName(layout.Name) ?? NormalizeName(layout.MatchingName);

        int slideCount = ctx.Presentation.Slides.Count;
        for (int i = 1; i <= slideCount; i++)
        {
            PowerPoint.Slide slide = ctx.Presentation.Slides[i];
            PowerPoint.CustomLayout? currentLayout = slide.CustomLayout;
            if (currentLayout is null)
            {
                continue;
            }

            if (IsSameComObject(currentLayout, layout)
                || IsSameComObject(currentLayout.Design, design)
                || IsSameComObject(slide.Design, design))
            {
                return true;
            }

            string? currentLayoutName = NormalizeName(currentLayout.Name) ?? NormalizeName(currentLayout.MatchingName);
            string? currentLayoutDesignName = NormalizeName(currentLayout.Design?.Name);
            string? slideDesignName = NormalizeName(slide.Design?.Name);

            if (string.Equals(currentLayoutName, targetLayoutName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(currentLayoutDesignName, targetDesignName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(slideDesignName, targetDesignName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsMasterUsedByAnySlide(PresentationContext ctx, PowerPoint.Design design)
    {
        string? targetDesignName = NormalizeName(design.Name);

        int slideCount = ctx.Presentation.Slides.Count;
        for (int i = 1; i <= slideCount; i++)
        {
            PowerPoint.Slide slide = ctx.Presentation.Slides[i];
            PowerPoint.CustomLayout? currentLayout = slide.CustomLayout;
            if (currentLayout is null)
            {
                continue;
            }

            if (IsSameComObject(slide.Design, design)
                || IsSameComObject(currentLayout.Design, design))
            {
                return true;
            }

            string? slideDesignName = NormalizeName(slide.Design?.Name);
            string? currentLayoutDesignName = NormalizeName(currentLayout.Design?.Name);
            if (string.Equals(slideDesignName, targetDesignName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(currentLayoutDesignName, targetDesignName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsSameComObject(object? left, object? right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        if (left is null || right is null)
        {
            return false;
        }

        try
        {
            return Marshal.GetIUnknownForObject(left) == Marshal.GetIUnknownForObject(right);
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static void InvokeComMethod(object target, string methodName)
    {
        if (!string.Equals(methodName, "Delete", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException($"Unsupported COM method '{methodName}'.");
        }

        // The PowerPoint interop RCWs expose Delete via COM dispatch rather than as a managed
        // member. Invoke it dynamically so the underlying PowerPoint COM object receives the call.
        dynamic dynamicTarget = target;
        dynamicTarget.Delete();
    }

    private static string? NormalizeName(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? GetMasterName(PowerPoint.Design design, PowerPoint.Master slideMaster)
    {
        string? name = TryGetString(design.Name);
        if (!string.IsNullOrWhiteSpace(name))
        {
            return name;
        }

        return TryGetString(slideMaster.Design?.Name);
    }

    private static string? GetLayoutName(PowerPoint.CustomLayout layout)
    {
        string? name = TryGetString(layout.Name);
        if (!string.IsNullOrWhiteSpace(name))
        {
            return name;
        }

        return TryGetString(layout.MatchingName);
    }

    private static string? TryGetString(object? value)
        => value is null ? null : value.ToString();

    private static MasterOperationResult ReadFont(PowerPoint.Shape placeholder)
    {
        dynamic? font = null;
        try
        {
            font = placeholder.TextFrame.TextRange.Font;
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
        finally
        {
            if (font != null)
            {
                ComUtilities.Release(ref font!);
            }
        }
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
        dynamic? font = null;
        try
        {
            font = placeholder.TextFrame.TextRange.Font;

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
        finally
        {
            if (font != null)
            {
                ComUtilities.Release(ref font!);
            }
        }
    }

    private static MasterOperationResult NotFound(string placeholderName)
        => new()
        {
            Success = false,
            ErrorMessage = $"The slide master does not have a '{placeholderName}' placeholder."
        };
}
