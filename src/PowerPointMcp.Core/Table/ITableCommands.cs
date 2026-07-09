using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <summary>
/// Table commands: add a table shape, read/write cell text, insert/delete rows and columns,
/// format cell fill and borders, and merge cells. Operates within an already-open
/// IPresentationBatch, targeting a specific slide and table shape by their 1-based indices.
/// </summary>
[ServiceCategory("table", "Table")]
[McpTool("table", Title = "Table Operations", Destructive = true, Category = "content",
    Description = "Add a table shape, read/write cell text, edit rows/columns, and format cells in an open presentation session.")]
public interface ITableCommands
{
    /// <summary>Adds a new table shape with the given number of rows/columns to a slide.</summary>
    TableOperationResult AddTable(IPresentationBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height);

    /// <summary>Sets the text of a table cell (1-based row/column).</summary>
    TableOperationResult SetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string text);

    /// <summary>Gets the text of a table cell (1-based row/column).</summary>
    TableOperationResult GetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column);

    /// <summary>
    /// Inserts a new row into the table. If <paramref name="beforeRow"/> is omitted, the row is
    /// appended as the last row. Returns the new total <c>rowCount</c>.
    /// </summary>
    TableOperationResult InsertRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeRow = null);

    /// <summary>Deletes the row at the given 1-based index. Returns the new total <c>rowCount</c>.</summary>
    TableOperationResult DeleteRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int row);

    /// <summary>
    /// Inserts a new column into the table. If <paramref name="beforeColumn"/> is omitted, the
    /// column is appended as the last column. Returns the new total <c>columnCount</c>.
    /// </summary>
    TableOperationResult InsertColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeColumn = null);

    /// <summary>Deletes the column at the given 1-based index. Returns the new total <c>columnCount</c>.</summary>
    TableOperationResult DeleteColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int column);

    /// <summary>Sets a table cell's fill to a solid RGB color.</summary>
    TableOperationResult SetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, byte red, byte green, byte blue);

    /// <summary>Gets a table cell's solid fill color.</summary>
    TableOperationResult GetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column);

    /// <summary>
    /// Sets one or more properties of a single border of a table cell. <paramref name="borderType"/>
    /// is a <c>PpBorderType</c> enum member name (<c>"ppBorderTop"</c>, <c>"ppBorderBottom"</c>,
    /// <c>"ppBorderLeft"</c>, <c>"ppBorderRight"</c>, <c>"ppBorderDiagonalDown"</c>, or
    /// <c>"ppBorderDiagonalUp"</c>). Any other parameter left null is unchanged; passing
    /// <paramref name="red"/>/<paramref name="green"/>/<paramref name="blue"/> together sets the
    /// border color; <paramref name="dashStyle"/> is an <c>MsoLineDashStyle</c> enum member name.
    /// </summary>
    TableOperationResult SetCellBorder(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int row,
        int column,
        string borderType,
        byte? red = null,
        byte? green = null,
        byte? blue = null,
        float? weight = null,
        string? dashStyle = null,
        bool? visible = null);

    /// <summary>Gets a single border's color, weight, dash style, and visibility for a table cell.</summary>
    TableOperationResult GetCellBorder(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string borderType);

    /// <summary>
    /// Merges the cell at (<paramref name="row"/>, <paramref name="column"/>) with the cell at
    /// (<paramref name="mergeToRow"/>, <paramref name="mergeToColumn"/>), producing a single
    /// merged cell. The two cells must be adjacent (in the same row or column).
    /// </summary>
    TableOperationResult MergeCells(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, int mergeToRow, int mergeToColumn);
}
