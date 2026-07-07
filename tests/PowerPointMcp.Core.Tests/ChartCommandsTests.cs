using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for chart commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Chart")]
public class ChartCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ChartCommands _commands = new();

    public ChartCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddChart_Bar_CreatesShape_WithExpectedCategoryAndSeriesCounts()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
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

    [Fact]
    public void AddChart_WithInvalidChartType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["A"];
        double[] values = [1d];
        var result = _commands.AddChart(batch, 1, "not-a-real-type", 0f, 0f, 100f, 100f, categories, "S", values);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
