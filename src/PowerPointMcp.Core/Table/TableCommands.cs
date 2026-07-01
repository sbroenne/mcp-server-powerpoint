using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <inheritdoc cref="ITableCommands"/>
public sealed class TableCommands : ITableCommands
{
    /// <inheritdoc/>
    public TableOperationResult AddTable(IPresentationBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            // Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height) — all parameters
            // are plain ints/floats defined on the PowerPoint interop's own Shapes interface,
            // no office.dll-typed enum involved (unlike AddShape/AddTextbox).
            slide.Shapes.AddTable(rows, columns, left, top, width, height);
            int newShapeIndex = (int)slide.Shapes.Count; // always appended — see Shape domain notes

            return new TableOperationResult
            {
                Success = true,
                ShapeIndex = newShapeIndex,
                RowCount = rows,
                ColumnCount = columns
            };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult SetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out dynamic? table);
            if (validation is not null) return validation;

            table!.Cell(row, column).Shape.TextFrame.TextRange.Text = text;

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, CellText = text };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult GetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out dynamic? table);
            if (validation is not null) return validation;

            string text = (string)table!.Cell(row, column).Shape.TextFrame.TextRange.Text;

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, CellText = text };
        });
    }

    private static TableOperationResult? ValidateTableShape(
        PresentationContext ctx, int slideIndex, int shapeIndex, int row, int column, out dynamic? table)
    {
        table = null;

        int slideCount = ctx.Presentation.Slides.Count;
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }

        dynamic slide = ctx.Presentation.Slides[slideIndex];
        int shapeCount = slide.Shapes.Count;
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }

        dynamic shape = slide.Shapes[shapeIndex];
        bool hasTable = (int)shape.HasTable != 0; // HasTable is MsoTriState-like tri-state int
        if (!hasTable)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} does not contain a table."
            };
        }

        int rowCount = (int)shape.Table.Rows.Count;
        int colCount = (int)shape.Table.Columns.Count;
        if (row < 1 || row > rowCount)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Row {row} is out of range. The table has {rowCount} row(s) (valid range: 1-{rowCount})."
            };
        }
        if (column < 1 || column > colCount)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Column {column} is out of range. The table has {colCount} column(s) (valid range: 1-{colCount})."
            };
        }

        table = shape.Table;
        return null;
    }
}
