using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Table;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for table commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Table")]
public class TableCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly TableCommands _commands = new();

    [Fact]
    public void AddTable_CreatesShapeWithExpectedDimensions()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.AddTable(batch, 1, rows: 3, columns: 2, left: 10f, top: 10f, width: 300f, height: 150f);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.ShapeIndex);
            Assert.Equal(3, result.RowCount);
            Assert.Equal(2, result.ColumnCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_ThenGetCellText_RoundTrips_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
                var setResult = _commands.SetCellText(batch, 1, 1, row: 1, column: 1, text: "Revenue");
                Assert.True(setResult.Success);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var getResult = _commands.GetCellText(reopened, 1, 1, row: 1, column: 1);
            Assert.True(getResult.Success);
            Assert.Equal("Revenue", getResult.CellText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetCellText_WithOutOfRangeRow_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            _commands.AddTable(batch, 1, 2, 2, 0f, 0f, 200f, 100f);
            var result = _commands.GetCellText(batch, 1, 1, row: 99, column: 1);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_OnShapeWithoutTable_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
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
        finally
        {
            File.Delete(path);
        }
    }
}
