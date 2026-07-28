using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <inheritdoc cref="ILayoutCommands"/>
public sealed class LayoutCommands : ILayoutCommands
{
    /// <inheritdoc/>
    public LayoutOperationResult SetLayout(IPresentationBatch batch, int slideIndex, string layoutName)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(layoutName);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            if (!Enum.TryParse<PowerPoint.PpSlideLayout>(layoutName, ignoreCase: true, out var layout))
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{layoutName}' is not a recognized PpSlideLayout name (e.g. 'ppLayoutBlank', 'ppLayoutTitleOnly', 'ppLayoutText')."
                };
            }

            ctx.Presentation.Slides[slideIndex].Layout = layout;

            return new LayoutOperationResult { Success = true, LayoutName = layout.ToString() };
        });
    }

    /// <inheritdoc/>
    public LayoutOperationResult GetLayout(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            PowerPoint.PpSlideLayout layout = ctx.Presentation.Slides[slideIndex].Layout;

            return new LayoutOperationResult { Success = true, LayoutName = layout.ToString() };
        });
    }

    /// <inheritdoc/>
    public LayoutOperationResult ListLayouts(IPresentationBatch batch, int masterIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (masterIndex < 1)
        {
            return new LayoutOperationResult
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
                return new LayoutOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Master index {masterIndex} is out of range. The presentation has {designCount} master(s) (valid range: 1-{designCount})."
                };
            }

            PowerPoint.Design design = designs[masterIndex];
            PowerPoint.Master slideMaster = design.SlideMaster;
            PowerPoint.CustomLayouts customLayouts = slideMaster.CustomLayouts;
            int layoutCount = customLayouts.Count;
            var layouts = new List<LayoutOperationResult.LayoutInventoryEntry>(layoutCount);

            for (int i = 1; i <= layoutCount; i++)
            {
                PowerPoint.CustomLayout layout = customLayouts[i];
                layouts.Add(new LayoutOperationResult.LayoutInventoryEntry
                {
                    LayoutIndex = i,
                    LayoutName = GetLayoutName(layout),
                    IsUsed = IsLayoutUsedByAnySlide(ctx, design, layout)
                });
            }

            return new LayoutOperationResult
            {
                Success = true,
                MasterIndex = masterIndex,
                MasterName = GetMasterName(design, slideMaster),
                Layouts = layouts
            };
        });
    }

    /// <inheritdoc/>
    public LayoutOperationResult DeleteLayout(IPresentationBatch batch, int masterIndex, int layoutIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (masterIndex < 1)
        {
            return new LayoutOperationResult
            {
                Success = false,
                ErrorMessage = "Master index must be 1 or greater."
            };
        }

        if (layoutIndex < 1)
        {
            return new LayoutOperationResult
            {
                Success = false,
                ErrorMessage = "Layout index must be 1 or greater."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            try
            {
                PowerPoint.Designs designs = ctx.Presentation.Designs;
                int designCount = designs.Count;
                if (masterIndex > designCount)
                {
                    return new LayoutOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Master index {masterIndex} is out of range. The presentation has {designCount} master(s) (valid range: 1-{designCount})."
                    };
                }

                PowerPoint.Design design = designs[masterIndex];
                PowerPoint.Master slideMaster = design.SlideMaster;
                PowerPoint.CustomLayouts customLayouts = slideMaster.CustomLayouts;
                int layoutCount = customLayouts.Count;
                if (layoutIndex > layoutCount)
                {
                    return new LayoutOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Layout index {layoutIndex} is out of range. The slide master has {layoutCount} layout(s) (valid range: 1-{layoutCount})."
                    };
                }

                PowerPoint.CustomLayout layout = customLayouts[layoutIndex];
                if (IsLayoutUsedByAnySlide(ctx, design, layout))
                {
                    return new LayoutOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Layout '{GetLayoutName(layout)}' is still used by one or more slides."
                    };
                }

                InvokeComMethod(layout, "Delete");
                return new LayoutOperationResult { Success = true };
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"DeleteLayout failed while processing masterIndex={masterIndex}, layoutIndex={layoutIndex}.", ex);
            }
        });
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

            if (IsSameComObject(currentLayout, layout))
            {
                return true;
            }

            string? currentLayoutName = NormalizeName(currentLayout.Name) ?? NormalizeName(currentLayout.MatchingName);
            string? currentLayoutDesignName = NormalizeName(currentLayout.Design?.Name);

            if (string.Equals(currentLayoutName, targetLayoutName, StringComparison.OrdinalIgnoreCase)
                && (IsSameComObject(currentLayout.Design, design)
                    || string.Equals(currentLayoutDesignName, targetDesignName, StringComparison.OrdinalIgnoreCase)))
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
        dynamic? dynamicTarget = null;
        try
        {
            dynamicTarget = target;
            dynamicTarget.Delete();
        }
        finally
        {
            ComUtilities.Release(ref dynamicTarget!);
        }
    }

    private static string? GetMasterName(PowerPoint.Design design, PowerPoint.Master slideMaster)
    {
        string? name = TryGetString(design.Name);
        if (!string.IsNullOrWhiteSpace(name))
        {
            return name;
        }

        return TryGetString(slideMaster.Design?.Name);
    }

    private static string? NormalizeName(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

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
}
