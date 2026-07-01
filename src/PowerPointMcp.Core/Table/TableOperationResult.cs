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
}
