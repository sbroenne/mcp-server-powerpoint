using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <summary>
/// Table commands: add a table shape and read/write cell text. Operates within an
/// already-open IPresentationBatch, targeting a specific slide and table shape by their
/// 1-based indices.
/// </summary>
public interface ITableCommands
{
    /// <summary>Adds a new table shape with the given number of rows/columns to a slide.</summary>
    TableOperationResult AddTable(IPresentationBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height);

    /// <summary>Sets the text of a table cell (1-based row/column).</summary>
    TableOperationResult SetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string text);

    /// <summary>Gets the text of a table cell (1-based row/column).</summary>
    TableOperationResult GetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column);
}
