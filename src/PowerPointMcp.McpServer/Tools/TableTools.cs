using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Table;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Table tools: add a table shape and read/write cell text.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="TableCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class TableTools
{
    private static readonly TableCommands Commands = new();

    /// <summary>Adds a new table shape with the given number of rows/columns to a slide.</summary>
    [McpServerTool(Name = "add_table")]
    [Description("Add a new table shape with the given number of rows/columns to a slide.")]
    public static string AddTable(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the table to.")] int slideIndex,
        [Description("Number of rows in the new table.")] int rows,
        [Description("Number of columns in the new table.")] int columns,
        [Description("Left position in points.")] float left,
        [Description("Top position in points.")] float top,
        [Description("Width in points.")] float width,
        [Description("Height in points.")] float height,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_table", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddTable(batch, slideIndex, rows, columns, left, top, width, height));
        });

    /// <summary>Sets the text of a table cell (1-based row/column).</summary>
    [McpServerTool(Name = "set_cell_text")]
    [Description("Set the text of a table cell (1-based row/column).")]
    public static string SetCellText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the table.")] int slideIndex,
        [Description("1-based index of the table shape.")] int shapeIndex,
        [Description("1-based row index of the cell.")] int row,
        [Description("1-based column index of the cell.")] int column,
        [Description("New text content for the cell.")] string text,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_cell_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetCellText(batch, slideIndex, shapeIndex, row, column, text));
        });

    /// <summary>Gets the text of a table cell (1-based row/column).</summary>
    [McpServerTool(Name = "get_cell_text")]
    [Description("Get the text of a table cell (1-based row/column).")]
    public static string GetCellText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the table.")] int slideIndex,
        [Description("1-based index of the table shape.")] int shapeIndex,
        [Description("1-based row index of the cell.")] int row,
        [Description("1-based column index of the cell.")] int column,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_cell_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetCellText(batch, slideIndex, shapeIndex, row, column));
        });

    private static string SerializeResult(TableOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            shapeIndex = result.ShapeIndex,
            rowCount = result.RowCount,
            columnCount = result.ColumnCount,
            cellText = result.CellText,
            isError = result.Success ? (bool?)null : true
        });
}
