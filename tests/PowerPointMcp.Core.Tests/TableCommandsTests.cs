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
}
