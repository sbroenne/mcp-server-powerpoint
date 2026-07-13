using System.Linq;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <inheritdoc cref="IChartCommands"/>
public sealed class ChartCommands : IChartCommands
{
    // Shapes.AddChart2's XlChartType parameter is defined in the Excel object model
    // (Microsoft.Office.Interop.Excel), which this project deliberately does not reference
    // (PowerPoint's chart feature embeds a mini Excel workbook for chart data, but the chart
    // *shape* itself only needs the raw XlChartType integer). Called late-bound via dynamic
    // with the raw int constants, consistent with the office.dll-avoidance pattern used
    // throughout this project.
    private const int XlColumnClustered = 51; // bar/column chart
    private const int XlLine = 4;
    private const int XlPie = 5;
    private const int MsoTrue = -1; // MsoTriState.msoTrue for Office.Core-backed tri-state members (for example Shape.HasChart).

    /// <inheritdoc/>
    public ChartOperationResult AddChart(IPresentationBatch batch, int slideIndex, string chartType, float left, float top, float width, float height, IReadOnlyList<string> categories, string seriesName, IReadOnlyList<double> values)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(chartType);
        ArgumentNullException.ThrowIfNull(categories);
        ArgumentNullException.ThrowIfNull(seriesName);
        ArgumentNullException.ThrowIfNull(values);

        int? xlChartType = chartType.ToLowerInvariant() switch
        {
            "bar" => XlColumnClustered,
            "line" => XlLine,
            "pie" => XlPie,
            _ => null
        };

        if (xlChartType is null)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Unknown chart type '{chartType}'. Supported values: bar, line, pie."
            };
        }

        if (categories.Count != values.Count)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Category count ({categories.Count}) must match value count ({values.Count})."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic? chartShape = null;
            try
            {
                // Shapes.AddChart2(Style, XlChartType, Left, Top, Width, Height, NewLayout) -> Shape
                // Reason: AddChart2 is not exposed on the strongly-typed PIA Shapes interface, so it
                // must be invoked via dynamic late binding.
                chartShape = ((dynamic)slide.Shapes).AddChart2(-1, xlChartType.Value, left, top, width, height, true);
                PowerPoint.Chart chart = chartShape.Chart;

                WriteChartData(chart, categories, seriesName, values);

                // Same NoPIA .Index late-binding quirk as Shape domain — use Shapes.Count instead.
                int newIndex = slide.Shapes.Count;

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = slide.Shapes.Count
                };
            }
            finally
            {
                if (chartShape != null)
                {
                    ComUtilities.Release(ref chartShape!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult GetChartData(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // NOTE: HasChart is MsoTriState (an int), not a C# bool — msoTrue = -1.
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            dynamic? seriesCollection = null;
            dynamic? firstSeries = null;
            try
            {
                // Chart writes (e.g. immediately preceding AddChart/ReplaceChartData calls) can
                // settle asynchronously in the embedded Excel workbook, so SeriesCollection.Count
                // can transiently read back as 0 right after a write. Retry until it's non-zero,
                // matching the same settling tolerance used by AddSeries/ReplaceChartData.
                int seriesCount = 0;
                for (int attempt = 1; attempt <= TransientReadRetryAttempts; attempt++)
                {
                    seriesCollection = RetryTransientChartRead(() => chart.SeriesCollection());
                    seriesCount = RetryTransientChartRead(() => (int)seriesCollection.Count);
                    if (seriesCount > 0 || attempt == TransientReadRetryAttempts)
                    {
                        break;
                    }

                    ComUtilities.Release(ref seriesCollection!);
                    seriesCollection = null;
                    System.Threading.Thread.Sleep(TransientReadRetryDelayMs);
                }

                int categoryCount = 0;
                if (seriesCount > 0)
                {
                    firstSeries = RetryTransientChartRead(() => seriesCollection.Item(1));
                    categoryCount = ReadNonEmptyXValues(firstSeries).Length;
                }

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    SeriesCount = seriesCount,
                    CategoryCount = categoryCount
                };
            }
            finally
            {
                if (firstSeries != null)
                {
                    ComUtilities.Release(ref firstSeries!);
                }

                if (seriesCollection != null)
                {
                    ComUtilities.Release(ref seriesCollection!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult AddSeries(IPresentationBatch batch, int slideIndex, int shapeIndex, string seriesName, IReadOnlyList<double> values)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(seriesName);
        ArgumentNullException.ThrowIfNull(values);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            dynamic? seriesCollection = null;
            dynamic? firstSeries = null;
            dynamic? newSeries = null;
            try
            {
                seriesCollection = RetryTransientChartRead(() => chart.SeriesCollection());
                int existingSeriesCount = RetryTransientChartRead(() => (int)seriesCollection.Count);

                if (existingSeriesCount == 0)
                {
                    return new ChartOperationResult
                    {
                        Success = false,
                        ErrorMessage = "The chart has no existing series to determine its category count from."
                    };
                }

                firstSeries = RetryTransientChartRead(() => seriesCollection.Item(1));
                Array existingXValues = ReadNonEmptyXValues(firstSeries);
                int categoryCount = existingXValues.Length;

                if (values.Count != categoryCount)
                {
                    return new ChartOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Value count ({values.Count}) must match the chart's existing category count ({categoryCount})."
                    };
                }

                // SeriesCollection is backed by Excel interop types that are not present in the
                // embedded PowerPoint PIA. Keep this boundary narrowly late-bound.
                newSeries = seriesCollection.NewSeries();
                newSeries.Values = values.ToArray();
                newSeries.XValues = existingXValues;
                newSeries.Name = seriesName;

                int newSeriesCount = (int)chart.SeriesCollection().Count;

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    SeriesCount = newSeriesCount,
                    CategoryCount = categoryCount
                };
            }
            finally
            {
                if (newSeries != null)
                {
                    ComUtilities.Release(ref newSeries!);
                }

                if (firstSeries != null)
                {
                    ComUtilities.Release(ref firstSeries!);
                }

                if (seriesCollection != null)
                {
                    ComUtilities.Release(ref seriesCollection!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult ReplaceChartData(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        IReadOnlyList<string> categories,
        IReadOnlyList<string> seriesNames,
        IReadOnlyList<double> seriesValues)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(categories);
        ArgumentNullException.ThrowIfNull(seriesNames);
        ArgumentNullException.ThrowIfNull(seriesValues);

        if (seriesNames.Count == 0)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = "At least one series name is required."
            };
        }

        int expectedValueCount = categories.Count * seriesNames.Count;
        if (seriesValues.Count != expectedValueCount)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"seriesValues length ({seriesValues.Count}) must equal categories.Count ({categories.Count}) * seriesNames.Count ({seriesNames.Count}) = {expectedValueCount}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;

            // Avoid the embedded Excel workbook and update its late-bound SeriesCollection
            // directly, because Excel interop types are not part of the embedded PowerPoint PIA.
            dynamic? seriesCollection = null;
            try
            {
                seriesCollection = RetryTransientChartRead(() => chart.SeriesCollection());
                int existingSeriesCount = RetryTransientChartRead(() => (int)seriesCollection.Count);
                for (int i = existingSeriesCount; i >= 1; i--)
                {
                    dynamic? existingSeries = null;
                    try
                    {
                        existingSeries = seriesCollection.Item(i);
                        existingSeries.Delete();
                    }
                    finally
                    {
                        if (existingSeries != null)
                        {
                            ComUtilities.Release(ref existingSeries!);
                        }
                    }
                }

                string[] categoriesArray = categories.ToArray();
                for (int s = 0; s < seriesNames.Count; s++)
                {
                    var valuesForSeries = new double[categories.Count];
                    for (int i = 0; i < categories.Count; i++)
                    {
                        valuesForSeries[i] = seriesValues[(s * categories.Count) + i];
                    }

                    dynamic? newSeries = null;
                    try
                    {
                        newSeries = seriesCollection.NewSeries();
                        newSeries.Values = valuesForSeries;
                        newSeries.XValues = categoriesArray;
                        newSeries.Name = seriesNames[s];
                    }
                    finally
                    {
                        if (newSeries != null)
                        {
                            ComUtilities.Release(ref newSeries!);
                        }
                    }
                }

                int newSeriesCount = (int)chart.SeriesCollection().Count;

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    SeriesCount = newSeriesCount,
                    CategoryCount = categories.Count
                };
            }
            finally
            {
                if (seriesCollection != null)
                {
                    ComUtilities.Release(ref seriesCollection!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string title)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(title);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Title = title,
                HasTitle = true
            };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult GetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            bool hasTitle = (bool)chart.HasTitle;
            string? title = hasTitle ? (string)chart.ChartTitle.Text : null;

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                HasTitle = hasTitle,
                Title = title
            };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType, string title)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(axisType);
        ArgumentNullException.ThrowIfNull(title);

        PowerPoint.XlAxisType? xlAxisType = ResolveAxisType(axisType);
        if (xlAxisType is null)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Unknown axis type '{axisType}'. Supported values: category, value."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            dynamic? axis = null;
            try
            {
                axis = chart.Axes(xlAxisType.Value);
                axis.HasTitle = true;
                axis.AxisTitle.Text = title;

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    AxisType = axisType,
                    Title = title,
                    HasTitle = true
                };
            }
            finally
            {
                if (axis != null)
                {
                    ComUtilities.Release(ref axis!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult GetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(axisType);

        PowerPoint.XlAxisType? xlAxisType = ResolveAxisType(axisType);
        if (xlAxisType is null)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Unknown axis type '{axisType}'. Supported values: category, value."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            dynamic? axis = null;
            try
            {
                axis = chart.Axes(xlAxisType.Value);
                bool hasTitle = (bool)axis.HasTitle;
                string? title = hasTitle ? (string)axis.AxisTitle.Text : null;

                return new ChartOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    AxisType = axisType,
                    HasTitle = hasTitle,
                    Title = title
                };
            }
            finally
            {
                if (axis != null)
                {
                    ComUtilities.Release(ref axis!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            chart.HasLegend = visible;

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                LegendVisible = visible
            };
        });
    }

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1822:Mark members as static",
        Justification = "Kept as an instance command method for command-class consistency.")]
    public ChartOperationResult GetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // Reason: read via dynamic late binding for consistent MsoTriState handling with the
            // rest of this file's chart-shape checks.
            if ((int)((dynamic)shape).HasChart != MsoTrue)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            PowerPoint.Chart chart = shape.Chart;
            bool visible = (bool)chart.HasLegend;

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                LegendVisible = visible
            };
        });
    }

    private static PowerPoint.XlAxisType? ResolveAxisType(string axisType) => axisType.ToLowerInvariant() switch
    {
        "category" => PowerPoint.XlAxisType.xlCategory,
        "value" => PowerPoint.XlAxisType.xlValue,
        _ => null
    };

    private static void WriteChartData(PowerPoint.Chart chart, IReadOnlyList<string> categories, string seriesName, IReadOnlyList<double> values)
    {
        // SeriesCollection is backed by Excel types that are not present in the embedded
        // PowerPoint PIA, so this is a deliberately narrow late-bound boundary.
        dynamic? seriesCollection = null;
        dynamic? newSeries = null;
        try
        {
            seriesCollection = RetryTransientChartRead(() => chart.SeriesCollection());
            int existingSeriesCount = RetryTransientChartRead(() => (int)seriesCollection.Count);
            for (int i = existingSeriesCount; i >= 1; i--)
            {
                dynamic? existingSeries = null;
                try
                {
                    existingSeries = seriesCollection.Item(i);
                    existingSeries.Delete();
                }
                finally
                {
                    if (existingSeries != null)
                    {
                        ComUtilities.Release(ref existingSeries!);
                    }
                }
            }

            newSeries = seriesCollection.NewSeries();
            newSeries.Values = values.ToArray();
            newSeries.XValues = categories.ToArray();
            newSeries.Name = seriesName;
        }
        finally
        {
            if (newSeries != null)
            {
                ComUtilities.Release(ref newSeries!);
            }

            if (seriesCollection != null)
            {
                ComUtilities.Release(ref seriesCollection!);
            }
        }
    }

    // Chart writes can settle asynchronously; a short bounded retry handles transient reads.
    // Bumped from 10x300ms (3s) to 20x300ms (6s) after observing the shorter window was
    // insufficient for SeriesCollection/XValues to settle under heavier COM/system load.
    private const int TransientReadRetryAttempts = 20;
    private const int TransientReadRetryDelayMs = 300;

    /// <summary>
    /// Runs <paramref name="read"/>, retrying on any <see cref="COMException"/> (not just
    /// RPC_E_DISCONNECTED) up to <see cref="TransientReadRetryAttempts"/> times. Intended only
    /// for idempotent read operations immediately following a chart write, where a genuinely
    /// broken COM state still fails after the bounded retry window instead of hanging.
    /// </summary>
    private static T RetryTransientChartRead<T>(Func<T> read)
    {
        Exception? lastError = null;
        for (int attempt = 1; attempt <= TransientReadRetryAttempts; attempt++)
        {
            try
            {
                return read();
            }
            catch (COMException ex)
            {
                lastError = ex;
                if (attempt == TransientReadRetryAttempts)
                {
                    throw new InvalidOperationException(
                        $"The chart's COM object model did not settle after {TransientReadRetryAttempts} attempts.",
                        ex);
                }
                System.Threading.Thread.Sleep(TransientReadRetryDelayMs);
            }
        }
        throw new InvalidOperationException("Unreachable.", lastError);
    }

    private static Array ReadNonEmptyXValues(dynamic series)
    {
        Array xValues = Array.Empty<object>();
        for (int attempt = 1; attempt <= TransientReadRetryAttempts; attempt++)
        {
            xValues = RetryTransientChartRead(() => (Array)series.XValues);
            if (xValues.Length > 0 || attempt == TransientReadRetryAttempts)
            {
                return xValues;
            }

            Thread.Sleep(TransientReadRetryDelayMs);
        }

        return xValues;
    }

    private static ChartOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }

    private static ChartOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new ChartOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }
        return null;
    }
}
