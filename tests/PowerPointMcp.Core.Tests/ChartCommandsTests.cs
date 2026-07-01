using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for chart commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Chart")]
public class ChartCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ChartCommands _commands = new();

    [Fact]
    public void AddChart_Bar_CreatesShape_WithExpectedCategoryAndSeriesCounts()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = ComInterop.Session.PresentationSession.BeginBatch(path);
            var categories = new[] { "Q1", "Q2", "Q3" };
            var values = new[] { 10d, 20d, 30d };

            var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);

            Assert.True(addResult.Success);
            Assert.Null(addResult.ErrorMessage);
            Assert.Equal(1, addResult.ShapeIndex);
            Assert.Equal(1, addResult.ShapeCount);

            var dataResult = _commands.GetChartData(batch, 1, addResult.ShapeIndex!.Value);

            Assert.True(dataResult.Success);
            Assert.Equal(3, dataResult.CategoryCount);
            Assert.Equal(1, dataResult.SeriesCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AddChart_WithInvalidChartType_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = ComInterop.Session.PresentationSession.BeginBatch(path);
            string[] categories = ["A"];
            double[] values = [1d];
            var result = _commands.AddChart(batch, 1, "not-a-real-type", 0f, 0f, 100f, 100f, categories, "S", values);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
