using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop.Session;

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
    private const int MsoTrue = -1; // MsoTriState.msoTrue — AddChart2's NewLayout param is MsoTriState, not bool.

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

            dynamic slide = ctx.Presentation.Slides[slideIndex];

            // Shapes.AddChart2(Style, XlChartType, Left, Top, Width, Height, NewLayout) -> Shape
            dynamic chartShape = slide.Shapes.AddChart2(-1, xlChartType.Value, left, top, width, height, MsoTrue);
            dynamic chart = chartShape.Chart;

            WriteChartData(chart, categories, seriesName, values);

            // Same NoPIA .Index late-binding quirk as Shape domain — use Shapes.Count instead.
            int newIndex = (int)slide.Shapes.Count;

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = (int)slide.Shapes.Count
            };
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            // NOTE: HasChart is MsoTriState (an int), not a C# bool — msoTrue = -1.
            if ((int)shape.HasChart != -1)
            {
                return new ChartOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a chart."
                };
            }

            dynamic chart = shape.Chart;
            dynamic seriesCollection = chart.SeriesCollection();
            int seriesCount = (int)seriesCollection.Count;
            int categoryCount = 0;
            if (seriesCount > 0)
            {
                dynamic firstSeries = seriesCollection.Item(1);
                dynamic xValues = firstSeries.XValues;
                categoryCount = (int)((Array)xValues).Length;
            }

            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                SeriesCount = seriesCount,
                CategoryCount = categoryCount
            };
        });
    }

    private static void WriteChartData(dynamic chart, IReadOnlyList<string> categories, string seriesName, IReadOnlyList<double> values)
    {
        // Modern (AddChart2) charts store their data in an embedded mini Excel workbook,
        // reachable via Chart.ChartData.Workbook. We write cells late-bound via dynamic to
        // avoid any dependency on the Excel interop assembly.
        dynamic chartData = chart.ChartData;
        chartData.Activate();

        // NOTE (discovered via real integration test, not assumed): immediately after
        // ChartData.Activate(), the embedded mini-Excel process that hosts the chart's data
        // workbook is still starting up out-of-process; accessing ChartData.Workbook right away
        // intermittently throws a generic COMException(0x80004005). Retry briefly until the
        // out-of-process workbook is ready.
        dynamic? workbook = null;
        Exception? lastError = null;
        for (int attempt = 0; attempt < 10 && workbook is null; attempt++)
        {
            try
            {
                workbook = chartData.Workbook;
            }
            catch (Exception ex)
            {
                lastError = ex;
                System.Threading.Thread.Sleep(200);
            }
        }
        if (workbook is null)
        {
            throw new InvalidOperationException("Timed out waiting for the chart's embedded data workbook to become available.", lastError);
        }

        dynamic worksheet = workbook.Worksheets[1];

        // NOTE (discovered via real integration test, not assumed): during a sustained
        // multi-hour PowerPoint session, the embedded chart-data workbook's out-of-process
        // Excel host can transiently drop its RPC connection mid-call, surfacing as
        // COMException(0x80010108 RPC_E_DISCONNECTED, "the object invoked has disconnected
        // from its clients") on the very NEXT dynamic dispatch (e.g. a Cells[...].Value2
        // write or a UsedRange access), even though Workbook itself resolved fine moments
        // earlier. This is transient — retrying the same call shortly after succeeds. Only
        // this specific HResult is retried; any other exception (bad argument, logic error)
        // propagates immediately without a retry.
        RetryOnDisconnect(() =>
        {
            worksheet.Cells[1, 1].Value2 = "Category";
            worksheet.Cells[1, 2].Value2 = seriesName;

            for (int i = 0; i < categories.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value2 = categories[i];
                worksheet.Cells[i + 2, 2].Value2 = values[i];
            }
        });

        // Clear any leftover default sample rows below our data.
        dynamic usedRange = null!;
        int usedRowCount = 0;
        RetryOnDisconnect(() =>
        {
            usedRange = worksheet.UsedRange;
            usedRowCount = (int)usedRange.Rows.Count;
        });
        int dataRowCount = categories.Count + 1; // + header row
        if (usedRowCount > dataRowCount)
        {
            // NOTE: build the range from an A1-style string address rather than
            // worksheet.Range[cellA, cellB] (two-Cells-object indexer) — the latter
            // intermittently throws ArgumentException ("Could not convert argument 0") when
            // both operands are themselves dynamic COM proxies from the embedded chart-data
            // workbook. A plain string address avoids the ambiguous dynamic-to-dynamic
            // indexer dispatch entirely.
            RetryOnDisconnect(() =>
            {
                dynamic extraRange = worksheet.Range[$"A{dataRowCount + 1}:B{usedRowCount}"];
                extraRange.ClearContents();
            });
        }

        // NOTE (discovered via real integration test, not assumed): calling
        // chart.SetSourceData(...) with a Range *object* (whether via `dynamic` DLR binding or
        // Type.InvokeMember) fails against the modern (AddChart2) chart engine — the Range lives
        // in the embedded chart-data workbook's process and PowerPoint's chart engine cannot
        // resolve it as a valid Source argument here (DISP_E_TYPEMISMATCH / "Could not convert
        // argument 0"). The modern chart engine instead expects Source as a plain sheet-qualified
        // A1-style STRING reference (e.g. "Sheet1!$A$1:$B$4"), which SetSourceData resolves
        // against the chart's own ChartData.Workbook internally. Using a string avoids passing a
        // cross-process COM object entirely.
        string worksheetName = worksheet.Name;
        string sourceAddress = $"{worksheetName}!$A$1:$B${dataRowCount}";
        object chartObj = chart;
        chartObj.GetType().InvokeMember(
            "SetSourceData",
            System.Reflection.BindingFlags.InvokeMethod,
            null,
            chartObj,
            [sourceAddress],
            System.Globalization.CultureInfo.InvariantCulture);

        // NOTE (discovered via real integration test, not assumed): Chart.SetSourceData commits
        // the chart data and often auto-closes the embedded chart-data Excel workbook/process as
        // part of that commit. A subsequent explicit Quit() on an already-closed workbook then
        // throws COMException(0x80010108 RPC_E_DISCONNECTED, "the object invoked has
        // disconnected from its clients"). Quit() is still attempted (some PowerPoint/Office
        // versions do NOT auto-close it), but failures are expected/best-effort here.
        try
        {
            workbook.Application.Quit();
        }
        catch (COMException)
        {
            // Embedded workbook was already closed by SetSourceData — nothing to do.
        }
    }

    // RPC_E_DISCONNECTED — thrown when a late-bound COM call reaches an object whose
    // out-of-process server (here, the embedded chart-data mini-Excel host) has dropped the
    // RPC connection. Bounded, fast-fail retry: a handful of short-backoff attempts is enough
    // to ride out a transient disconnect, while a genuinely broken COM state still fails
    // quickly instead of hanging.
    private const int RpcEDisconnected = unchecked((int)0x80010108);
    private const int DisconnectRetryAttempts = 4;
    private const int DisconnectRetryDelayMs = 150;

    /// <summary>
    /// Runs <paramref name="action"/>, retrying only on COMException(RPC_E_DISCONNECTED)
    /// (see <see cref="RpcEDisconnected"/>). Any other exception — including a bad argument
    /// or logic error — propagates immediately without a retry, since only the transient
    /// disconnect is safe to silently retry.
    /// </summary>
    private static void RetryOnDisconnect(Action action)
    {
        for (int attempt = 1; attempt <= DisconnectRetryAttempts; attempt++)
        {
            try
            {
                action();
                return;
            }
            catch (COMException ex) when (ex.HResult == RpcEDisconnected)
            {
                if (attempt == DisconnectRetryAttempts)
                {
                    throw new InvalidOperationException(
                        $"The chart's embedded data workbook disconnected from its COM client (RPC_E_DISCONNECTED) after {DisconnectRetryAttempts} attempts while writing chart data.",
                        ex);
                }
                System.Threading.Thread.Sleep(DisconnectRetryDelayMs);
            }
        }
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
