using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Table;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for table commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Table")]
public class TableCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly TableCommands _commands = new();

    public TableCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddTable_CreatesShapeWithExpectedDimensions()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddTable(batch, 1, rows: 3, columns: 2, left: 10f, top: 10f, width: 300f, height: 150f);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public void SetCellText_ThenGetCellText_RoundTrips_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
        var setResult = _commands.SetCellText(batch, 1, 1, row: 1, column: 1, text: "Revenue");
        Assert.True(setResult.Success);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var getResult = _commands.GetCellText(batch, 1, 1, row: 1, column: 1);
        Assert.True(getResult.Success);
        Assert.Equal("Revenue", getResult.CellText);
    }

    [Fact]
    public void GetCellText_WithOutOfRangeRow_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
        var result = _commands.GetCellText(batch, 1, 1, row: 99, column: 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetCellText_OnShapeWithoutTable_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        // The default first slide's placeholder shapes have no table.
        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 50f, 50f);
            return 0;
        });

        var result = _commands.SetCellText(batch, 1, 1, row: 1, column: 1, text: "x");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void InsertRow_Appended_IncreasesRowCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.InsertRow(batch, 1, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public void InsertRow_WithBeforeRow_InsertsAtPosition()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
        _commands.SetCellText(batch, 1, 1, row: 1, column: 1, text: "First");
        _commands.SetCellText(batch, 1, 1, row: 2, column: 1, text: "Second");

        var result = _commands.InsertRow(batch, 1, 1, beforeRow: 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(3, result.RowCount);
        Assert.Equal("First", _commands.GetCellText(batch, 1, 1, row: 1, column: 1).CellText);
        Assert.Equal("Second", _commands.GetCellText(batch, 1, 1, row: 3, column: 1).CellText);
    }

    [Fact]
    public void InsertRow_WithInvalidBeforeRow_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.InsertRow(batch, 1, 1, beforeRow: 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void DeleteRow_DecreasesRowCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 3, 2, 0f, 0f, 200f, 100f);

        var result = _commands.DeleteRow(batch, 1, 1, row: 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.RowCount);
    }

    [Fact]
    public void DeleteRow_WithInvalidRow_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.DeleteRow(batch, 1, 1, row: 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void InsertColumn_Appended_IncreasesColumnCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.InsertColumn(batch, 1, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
    }

    [Fact]
    public void InsertColumn_WithInvalidBeforeColumn_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.InsertColumn(batch, 1, 1, beforeColumn: 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void DeleteColumn_DecreasesColumnCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 3, 0f, 0f, 200f, 100f);

        var result = _commands.DeleteColumn(batch, 1, 1, column: 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public void DeleteColumn_WithInvalidColumn_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.DeleteColumn(batch, 1, 1, column: 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetCellFill_AndGetCellFill_RoundTripsColor()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var setResult = _commands.SetCellFill(batch, 1, 1, row: 1, column: 1, red: 255, green: 0, blue: 0);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);

        var getResult = _commands.GetCellFill(batch, 1, 1, row: 1, column: 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(255, getResult.ColorRgb);
    }

    [Fact]
    public void SetCellFill_WithInvalidRow_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.SetCellFill(batch, 1, 1, row: 99, column: 1, red: 255, green: 0, blue: 0);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetCellBorder_AndGetCellBorder_RoundTripsColorWeightDashStyleAndVisibility()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var setResult = _commands.SetCellBorder(batch, 1, 1, row: 1, column: 1, borderType: "ppBorderBottom",
            red: 0, green: 255, blue: 0, weight: 2f, dashStyle: "msoLineDash", visible: true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppBorderBottom", setResult.BorderType);
        Assert.Equal(65280, setResult.ColorRgb);
        Assert.Equal(2f, setResult.LineWeight);
        Assert.Equal("msoLineDash", setResult.DashStyleName);
        Assert.True(setResult.Visible);

        var getResult = _commands.GetCellBorder(batch, 1, 1, row: 1, column: 1, borderType: "ppBorderBottom");
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(65280, getResult.ColorRgb);
        Assert.Equal(2f, getResult.LineWeight);
        Assert.Equal("msoLineDash", getResult.DashStyleName);
        Assert.True(getResult.Visible);
    }

    [Fact]
    public void SetCellBorder_WithUnrecognizedBorderType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.SetCellBorder(batch, 1, 1, row: 1, column: 1, borderType: "ppBorderDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetCellBorder_WithUnrecognizedDashStyle_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.SetCellBorder(batch, 1, 1, row: 1, column: 1, borderType: "ppBorderTop", dashStyle: "msoLineDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void MergeCells_ReducesCellCountWithoutChangingGridDimensions()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
        _commands.SetCellText(batch, 1, 1, row: 1, column: 1, text: "Merged");

        var result = _commands.MergeCells(batch, 1, 1, row: 1, column: 1, mergeToRow: 1, mergeToColumn: 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);

        // After merging, reading the merged cell's text (from either original coordinate) should
        // still return the pre-merge content.
        var textResult = _commands.GetCellText(batch, 1, 1, row: 1, column: 1);
        Assert.True(textResult.Success, textResult.ErrorMessage);
        Assert.Equal("Merged", textResult.CellText);
    }

    [Fact]
    public void MergeCells_WithInvalidMergeToCoordinates_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);

        var result = _commands.MergeCells(batch, 1, 1, row: 1, column: 1, mergeToRow: 99, mergeToColumn: 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
