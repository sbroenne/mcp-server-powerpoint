using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.McpServer.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Chart tools: add a native chart shape with categories/series data, and read chart data
/// dimensions back.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="ChartCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows. Array-typed Core parameters
/// (<c>IReadOnlyList&lt;string&gt;</c>/<c>IReadOnlyList&lt;double&gt;</c>) are exposed as
/// <c>string[]</c>/<c>double[]</c> tool parameters — the MCP SDK (1.3.0) serializes these into a
/// JSON array schema and the arrays are passed straight through to Core.
/// </remarks>
[McpServerToolType]
public static class ChartTools
{
    private static readonly ChartCommands Commands = new();

    /// <summary>
    /// Adds a native chart shape (bar, line, or pie) to the given slide with categories and a
    /// single data series.
    /// </summary>
    [McpServerTool(Name = "add_chart")]
    [Description("Add a native chart shape (bar, line, or pie) to the given slide with categories and a single data series.")]
    public static string AddChart(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the chart to.")] int slideIndex,
        [Description("Chart type: \"bar\", \"line\", or \"pie\".")] string chartType,
        [Description("Left position in points.")] float left,
        [Description("Top position in points.")] float top,
        [Description("Width in points.")] float width,
        [Description("Height in points.")] float height,
        [Description("Category labels (x-axis / pie slice labels).")] string[] categories,
        [Description("Name of the single data series.")] string seriesName,
        [Description("Data values, one per category.")] double[] values,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_chart", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddChart(batch, slideIndex, chartType, left, top, width, height, categories, seriesName, values));
        });

    /// <summary>Gets the category and series counts of an existing chart shape's data.</summary>
    [McpServerTool(Name = "get_chart_data")]
    [Description("Get the category and series counts of an existing chart shape's data.")]
    public static string GetChartData(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the chart.")] int slideIndex,
        [Description("1-based index of the chart shape.")] int shapeIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_chart_data", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetChartData(batch, slideIndex, shapeIndex));
        });

    private static string SerializeResult(ChartOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            shapeIndex = result.ShapeIndex,
            shapeCount = result.ShapeCount,
            categoryCount = result.CategoryCount,
            seriesCount = result.SeriesCount,
            isError = result.Success ? (bool?)null : true
        });
}
