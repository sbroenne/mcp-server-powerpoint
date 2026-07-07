using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>
/// Chart lifecycle and data operations.
/// </summary>
[ServiceCategory("chart", "Chart")]
[McpTool("chart", Title = "Chart Operations", Destructive = true, Category = "content",
    Description = "Add a native chart shape and read chart data in an open presentation session.")]
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
}
