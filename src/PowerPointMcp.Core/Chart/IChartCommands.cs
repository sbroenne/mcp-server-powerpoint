using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>
/// Chart lifecycle and data operations.
/// </summary>
[ServiceCategory("chart", "Chart")]
[McpTool("chart", Title = "Chart Operations", Destructive = true, Category = "content",
    Description = "Add a native chart shape, add data series, set titles/legend, and read chart data in an open presentation session.")]
public interface IChartCommands
{
    /// <summary>
    /// Adds a native chart shape (bar, line, or pie) to the given slide with categories and a
    /// single data series, and returns the new shape's index.
    /// </summary>
    /// <param name="batch">The active presentation batch.</param>
    /// <param name="slideIndex">1-based slide index.</param>
    /// <param name="chartType">Chart type: "bar", "line", or "pie".</param>
    /// <param name="left">Left position in points.</param>
    /// <param name="top">Top position in points.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="categories">Category labels (x-axis / pie slice labels).</param>
    /// <param name="seriesName">Name of the single data series.</param>
    /// <param name="values">Data values, one per category.</param>
    ChartOperationResult AddChart(IPresentationBatch batch, int slideIndex, string chartType, float left, float top, float width, float height, IReadOnlyList<string> categories, string seriesName, IReadOnlyList<double> values);

    /// <summary>
    /// Gets the category and series counts of an existing chart shape's data.
    /// </summary>
    ChartOperationResult GetChartData(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Adds another data series to an existing chart (created via <see cref="AddChart"/>),
    /// producing a multi-series chart. <paramref name="values"/> must have the same count as the
    /// chart's existing categories. Returns the new total <c>seriesCount</c>.
    /// </summary>
    ChartOperationResult AddSeries(IPresentationBatch batch, int slideIndex, int shapeIndex, string seriesName, IReadOnlyList<double> values);

    /// <summary>Sets the chart's main title text (and turns the title on).</summary>
    ChartOperationResult SetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string title);

    /// <summary>Gets the chart's main title text and whether it currently has a title (<c>hasTitle</c>).</summary>
    ChartOperationResult GetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets an axis title (and turns that axis's title on). <paramref name="axisType"/> is
    /// <c>"category"</c> or <c>"value"</c>.
    /// </summary>
    ChartOperationResult SetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType, string title);

    /// <summary>
    /// Gets an axis's title text and whether it currently has a title (<c>hasTitle</c>).
    /// <paramref name="axisType"/> is <c>"category"</c> or <c>"value"</c>.
    /// </summary>
    ChartOperationResult GetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType);

    /// <summary>Shows or hides the chart's legend.</summary>
    ChartOperationResult SetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible);

    /// <summary>Gets whether the chart's legend is visible.</summary>
    ChartOperationResult GetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Replaces ALL of an existing chart's data (categories and every series) in one call — the
    /// chart's previous categories and series are discarded and replaced wholesale. Unlike
    /// <see cref="AddSeries"/> (which appends one series to the existing categories),
    /// <paramref name="categories"/> here can also change the category count/labels.
    /// <paramref name="seriesValues"/> is a single flat list laid out series-major: all values
    /// for <c>seriesNames[0]</c> (one per category, in category order), then all values for
    /// <c>seriesNames[1]</c>, and so on. Its length must equal
    /// <c>categories.Count * seriesNames.Count</c>.
    /// </summary>
    ChartOperationResult ReplaceChartData(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        IReadOnlyList<string> categories,
        IReadOnlyList<string> seriesNames,
        IReadOnlyList<double> seriesValues);
}
