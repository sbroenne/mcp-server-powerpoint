using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <inheritdoc cref="ITableCommands"/>
public sealed class TableCommands : ITableCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // PpBorderType, verified against learn.microsoft.com/office/vba/api/powerpoint.ppbordertype.
    private static readonly Dictionary<string, PowerPoint.PpBorderType> BorderTypes = new()
    {
        ["ppBorderTop"] = PowerPoint.PpBorderType.ppBorderTop,
        ["ppBorderLeft"] = PowerPoint.PpBorderType.ppBorderLeft,
        ["ppBorderBottom"] = PowerPoint.PpBorderType.ppBorderBottom,
        ["ppBorderRight"] = PowerPoint.PpBorderType.ppBorderRight,
        ["ppBorderDiagonalDown"] = PowerPoint.PpBorderType.ppBorderDiagonalDown,
        ["ppBorderDiagonalUp"] = PowerPoint.PpBorderType.ppBorderDiagonalUp,
    };

    // MsoLineDashStyle, verified against learn.microsoft.com/office/vba/api/office.msolinedashstyle
    // (same curated subset used by ShapeCommands.SetLine/GetLine).
    private static readonly Dictionary<string, int> LineDashStyles = new()
    {
        ["msoLineSolid"] = 1,
        ["msoLineSquareDot"] = 2,
        ["msoLineRoundDot"] = 3,
        ["msoLineDash"] = 4,
        ["msoLineDashDot"] = 5,
        ["msoLineDashDotDot"] = 6,
        ["msoLineLongDash"] = 7,
        ["msoLineLongDashDot"] = 8,
    };

    private static readonly Dictionary<int, string> LineDashStylesByValue =
        LineDashStyles.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            // Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height) — all parameters
            // are plain ints/floats defined on the PowerPoint interop's own Shapes interface,
            // no office.dll-typed enum involved (unlike AddShape/AddTextbox).
            slide.Shapes.AddTable(rows, columns, left, top, width, height);
            int newShapeIndex = slide.Shapes.Count; // always appended — see Shape domain notes

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
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
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
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            string text = (string)table!.Cell(row, column).Shape.TextFrame.TextRange.Text;

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, CellText = text };
        });
    }

    private static TableOperationResult? ValidateTableShape(
        PresentationContext ctx, int slideIndex, int shapeIndex, int row, int column, out PowerPoint.Table? table)
    {
        var validation = ValidateTableShapeOnly(ctx, slideIndex, shapeIndex, out table);
        if (validation is not null) return validation;

        int rowCount = table!.Rows.Count;
        int colCount = table.Columns.Count;
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

        return null;
    }

    private static TableOperationResult? ValidateTableShapeOnly(
        PresentationContext ctx, int slideIndex, int shapeIndex, out PowerPoint.Table? table)
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

        PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
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
        bool hasTable = (int)shape.HasTable == MsoTrue;
        if (!hasTable)
        {
            return new TableOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} does not contain a table."
            };
        }

        table = (PowerPoint.Table)shape.Table;
        return null;
    }

    /// <inheritdoc/>
    public TableOperationResult InsertRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeRow = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShapeOnly(ctx, slideIndex, shapeIndex, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int rowCount = table!.Rows.Count;
            if (beforeRow is not null && (beforeRow < 1 || beforeRow > rowCount))
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"BeforeRow {beforeRow} is out of range. The table has {rowCount} row(s) (valid range: 1-{rowCount})."
                };
            }

            if (beforeRow is null)
            {
                table.Rows.Add();
            }
            else
            {
                table.Rows.Add(beforeRow.Value);
            }

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = table.Rows.Count, ColumnCount = table.Columns.Count };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult DeleteRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int row)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShapeOnly(ctx, slideIndex, shapeIndex, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int rowCount = table!.Rows.Count;
            if (row < 1 || row > rowCount)
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Row {row} is out of range. The table has {rowCount} row(s) (valid range: 1-{rowCount})."
                };
            }

            table.Rows[row].Delete();

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = table.Rows.Count, ColumnCount = table.Columns.Count };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult InsertColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeColumn = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShapeOnly(ctx, slideIndex, shapeIndex, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int colCount = table!.Columns.Count;
            if (beforeColumn is not null && (beforeColumn < 1 || beforeColumn > colCount))
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"BeforeColumn {beforeColumn} is out of range. The table has {colCount} column(s) (valid range: 1-{colCount})."
                };
            }

            if (beforeColumn is null)
            {
                table.Columns.Add();
            }
            else
            {
                table.Columns.Add(beforeColumn.Value);
            }

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = table.Rows.Count, ColumnCount = table.Columns.Count };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult DeleteColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int column)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShapeOnly(ctx, slideIndex, shapeIndex, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int colCount = table!.Columns.Count;
            if (column < 1 || column > colCount)
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Column {column} is out of range. The table has {colCount} column(s) (valid range: 1-{colCount})."
                };
            }

            table.Columns[column].Delete();

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = table.Rows.Count, ColumnCount = table.Columns.Count };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult SetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, byte red, byte green, byte blue)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int rgb = red + (green << 8) + (blue << 16);
            PowerPoint.Shape cellShape = table!.Cell(row, column).Shape;
            cellShape.Fill.Solid();
            cellShape.Fill.ForeColor.RGB = rgb;

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult GetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            PowerPoint.Shape cellShape = table!.Cell(row, column).Shape;
            int rgb = cellShape.Fill.ForeColor.RGB;

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult SetCellBorder(
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
        bool? visible = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(borderType);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            if (!BorderTypes.TryGetValue(borderType, out var borderTypeValue))
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{borderType}' is not a recognized PpBorderType name (must be 'ppBorderTop', 'ppBorderBottom', 'ppBorderLeft', 'ppBorderRight', 'ppBorderDiagonalDown', or 'ppBorderDiagonalUp')."
                };
            }

            int? dashStyleValue = null;
            if (dashStyle is not null)
            {
                if (!LineDashStyles.TryGetValue(dashStyle, out var resolvedDashStyle))
                {
                    return new TableOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"'{dashStyle}' is not a recognized MsoLineDashStyle name (e.g. 'msoLineSolid', 'msoLineDash', 'msoLineDashDot')."
                    };
                }
                dashStyleValue = resolvedDashStyle;
            }

            PowerPoint.LineFormat border = table!.Cell(row, column).Borders[borderTypeValue];

            if (red is not null || green is not null || blue is not null)
            {
                int rgb = (red ?? 0) + ((green ?? 0) << 8) + ((blue ?? 0) << 16);
                border.ForeColor.RGB = rgb;
            }

            if (weight is not null)
            {
                border.Weight = weight.Value;
            }

            if (dashStyleValue is not null)
            {
                dynamic dynBorder = border;
                dynBorder.DashStyle = dashStyleValue.Value;
            }

            if (visible is not null)
            {
                dynamic dynBorder = border;
                dynBorder.Visible = visible.Value ? MsoTrue : MsoFalse;
            }

            return ReadBorder(border, shapeIndex, borderType);
        });
    }

    /// <inheritdoc/>
    public TableOperationResult GetCellBorder(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string borderType)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(borderType);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            if (!BorderTypes.TryGetValue(borderType, out var borderTypeValue))
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{borderType}' is not a recognized PpBorderType name (must be 'ppBorderTop', 'ppBorderBottom', 'ppBorderLeft', 'ppBorderRight', 'ppBorderDiagonalDown', or 'ppBorderDiagonalUp')."
                };
            }

            PowerPoint.LineFormat border = table!.Cell(row, column).Borders[borderTypeValue];
            return ReadBorder(border, shapeIndex, borderType);
        });
    }

    /// <inheritdoc/>
    public TableOperationResult MergeCells(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, int mergeToRow, int mergeToColumn)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateTableShape(ctx, slideIndex, shapeIndex, row, column, out PowerPoint.Table? table);
            if (validation is not null) return validation;

            int rowCount = table!.Rows.Count;
            int colCount = table.Columns.Count;
            if (mergeToRow < 1 || mergeToRow > rowCount)
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"MergeTo row {mergeToRow} is out of range. The table has {rowCount} row(s) (valid range: 1-{rowCount})."
                };
            }
            if (mergeToColumn < 1 || mergeToColumn > colCount)
            {
                return new TableOperationResult
                {
                    Success = false,
                    ErrorMessage = $"MergeTo column {mergeToColumn} is out of range. The table has {colCount} column(s) (valid range: 1-{colCount})."
                };
            }

            table!.Cell(row, column).Merge(table.Cell(mergeToRow, mergeToColumn));

            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = rowCount, ColumnCount = colCount };
        });
    }

    private static TableOperationResult ReadBorder(PowerPoint.LineFormat border, int shapeIndex, string borderType)
    {
        int rgb = border.ForeColor.RGB;
        float weight = border.Weight;
        dynamic dynBorder = border;
        int dashStyleValue = (int)dynBorder.DashStyle;
        bool visible = (int)dynBorder.Visible == MsoTrue;

        LineDashStylesByValue.TryGetValue(dashStyleValue, out var dashStyleName);

        return new TableOperationResult
        {
            Success = true,
            ShapeIndex = shapeIndex,
            BorderType = borderType,
            ColorRgb = rgb,
            LineWeight = weight,
            DashStyleName = dashStyleName,
            Visible = visible
        };
    }
}
