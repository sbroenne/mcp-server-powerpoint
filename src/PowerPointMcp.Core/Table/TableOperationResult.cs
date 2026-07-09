namespace Sbroenne.PowerPointMcp.Core.Table;

/// <summary>
/// Result of a table operation (add table, set/get cell text).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class TableOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the shape holding the table, if applicable.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Number of rows in the table, if applicable.</summary>
    public int? RowCount { get; init; }

    /// <summary>Number of columns in the table, if applicable.</summary>
    public int? ColumnCount { get; init; }

    /// <summary>Cell text content, for GetCellText or after SetCellText.</summary>
    public string? CellText { get; init; }

    /// <summary>Fill or border color as an RGB integer (0xBBGGRR, PowerPoint's native color order), if applicable.</summary>
    public int? ColorRgb { get; init; }

    /// <summary>The PpBorderType name of the border acted on, if applicable.</summary>
    public string? BorderType { get; init; }

    /// <summary>Border weight in points, if applicable.</summary>
    public float? LineWeight { get; init; }

    /// <summary>The MsoLineDashStyle name of the border, if applicable.</summary>
    public string? DashStyleName { get; init; }

    /// <summary>Whether the border is visible, if applicable.</summary>
    public bool? Visible { get; init; }
}
