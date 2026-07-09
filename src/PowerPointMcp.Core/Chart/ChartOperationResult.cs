namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>
/// Result of a chart-domain operation.
/// </summary>
public sealed class ChartOperationResult
{
    /// <summary>True if the operation succeeded. Never true when <see cref="ErrorMessage"/> is set.</summary>
    public bool Success { get; init; }

    /// <summary>Human-readable error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the affected chart shape within the slide's Shapes collection.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Total shape count on the slide after the operation.</summary>
    public int? ShapeCount { get; init; }

    /// <summary>Number of categories in the chart's data (for data queries).</summary>
    public int? CategoryCount { get; init; }

    /// <summary>Number of series in the chart's data (for data queries).</summary>
    public int? SeriesCount { get; init; }

    /// <summary>Chart or axis title text, for title queries/updates.</summary>
    public string? Title { get; init; }

    /// <summary>Whether the chart or axis currently has a title.</summary>
    public bool? HasTitle { get; init; }

    /// <summary>Axis type the operation acted on ("category" or "value"), echoed back.</summary>
    public string? AxisType { get; init; }

    /// <summary>Whether the chart's legend is currently visible.</summary>
    public bool? LegendVisible { get; init; }
}
