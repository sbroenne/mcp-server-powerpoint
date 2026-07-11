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

    [Fact]
    public void AddSeries_ToExistingChart_IncreasesSeriesCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2", "Q3"];
        double[] values = [10d, 20d, 30d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var seriesResult = _commands.AddSeries(batch, 1, shapeIndex, "Costs", [5d, 12d, 18d]);

        Assert.True(seriesResult.Success, seriesResult.ErrorMessage);
        Assert.Equal(2, seriesResult.SeriesCount);
        Assert.Equal(3, seriesResult.CategoryCount);

        var dataResult = _commands.GetChartData(batch, 1, shapeIndex);
        Assert.True(dataResult.Success, dataResult.ErrorMessage);
        Assert.Equal(2, dataResult.SeriesCount);
        Assert.Equal(3, dataResult.CategoryCount);
    }

    [Fact]
    public void AddSeries_WithMismatchedValueCount_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2", "Q3"];
        double[] values = [10d, 20d, 30d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var result = _commands.AddSeries(batch, 1, shapeIndex, "Costs", [5d, 12d]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddSeries_OnShapeWithoutChart_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 50f, 50f);
            return 0;
        });

        var result = _commands.AddSeries(batch, 1, 1, "Costs", [1d]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ReplaceChartData_WithNewCategoriesAndMultipleSeries_ReplacesAllData()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] initialCategories = ["Q1", "Q2", "Q3"];
        double[] initialValues = [10d, 20d, 30d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, initialCategories, "Revenue", initialValues);
        int shapeIndex = addResult.ShapeIndex!.Value;

        // Replace with a different category count and two series (series-major flat layout:
        // all 4 "Revenue" values, then all 4 "Costs" values).
        string[] newCategories = ["Jan", "Feb", "Mar", "Apr"];
        string[] seriesNames = ["Revenue", "Costs"];
        double[] seriesValues = [100d, 200d, 300d, 400d, 50d, 60d, 70d, 80d];

        var replaceResult = _commands.ReplaceChartData(batch, 1, shapeIndex, newCategories, seriesNames, seriesValues);

        Assert.True(replaceResult.Success, replaceResult.ErrorMessage);
        Assert.Equal(4, replaceResult.CategoryCount);
        Assert.Equal(2, replaceResult.SeriesCount);

        var dataResult = _commands.GetChartData(batch, 1, shapeIndex);
        Assert.True(dataResult.Success, dataResult.ErrorMessage);
        Assert.Equal(4, dataResult.CategoryCount);
        Assert.Equal(2, dataResult.SeriesCount);
    }

    [Fact]
    public void ReplaceChartData_WithMismatchedValueCount_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2", "Q3"];
        double[] values = [10d, 20d, 30d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        // 2 categories * 2 series = 4 values expected, only 3 given.
        var result = _commands.ReplaceChartData(batch, 1, shapeIndex, ["A", "B"], ["S1", "S2"], [1d, 2d, 3d]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ReplaceChartData_OnShapeWithoutChart_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 50f, 50f);
            return 0;
        });

        var result = _commands.ReplaceChartData(batch, 1, 1, ["A"], ["S1"], [1d]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetChartTitle_ThenGetChartTitle_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2"];
        double[] values = [1d, 2d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var setResult = _commands.SetChartTitle(batch, 1, shapeIndex, "Quarterly Revenue");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.HasTitle);
        Assert.Equal("Quarterly Revenue", setResult.Title);

        var getResult = _commands.GetChartTitle(batch, 1, shapeIndex);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.HasTitle);
        Assert.Equal("Quarterly Revenue", getResult.Title);
    }

    [Fact]
    public void GetChartTitle_OnShapeWithoutChart_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 50f, 50f);
            return 0;
        });

        var result = _commands.GetChartTitle(batch, 1, 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAxisTitle_ThenGetAxisTitle_RoundTrips_ForCategoryAndValueAxes()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2"];
        double[] values = [1d, 2d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var setCategoryResult = _commands.SetAxisTitle(batch, 1, shapeIndex, "category", "Quarter");
        Assert.True(setCategoryResult.Success, setCategoryResult.ErrorMessage);
        Assert.Equal("category", setCategoryResult.AxisType);
        Assert.Equal("Quarter", setCategoryResult.Title);

        var getCategoryResult = _commands.GetAxisTitle(batch, 1, shapeIndex, "category");
        Assert.True(getCategoryResult.Success, getCategoryResult.ErrorMessage);
        Assert.Equal("Quarter", getCategoryResult.Title);

        var setValueResult = _commands.SetAxisTitle(batch, 1, shapeIndex, "value", "Dollars ($M)");
        Assert.True(setValueResult.Success, setValueResult.ErrorMessage);
        Assert.Equal("value", setValueResult.AxisType);

        var getValueResult = _commands.GetAxisTitle(batch, 1, shapeIndex, "value");
        Assert.True(getValueResult.Success, getValueResult.ErrorMessage);
        Assert.Equal("Dollars ($M)", getValueResult.Title);
    }

    [Fact]
    public void SetAxisTitle_WithUnrecognizedAxisType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1"];
        double[] values = [1d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var result = _commands.SetAxisTitle(batch, 1, shapeIndex, "not-a-real-axis", "x");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetLegendVisibility_ThenGetLegendVisibility_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string[] categories = ["Q1", "Q2"];
        double[] values = [1d, 2d];
        var addResult = _commands.AddChart(batch, 1, "bar", 50f, 50f, 400f, 300f, categories, "Revenue", values);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var hideResult = _commands.SetLegendVisibility(batch, 1, shapeIndex, visible: false);
        Assert.True(hideResult.Success, hideResult.ErrorMessage);
        Assert.False(hideResult.LegendVisible);

        var getHiddenResult = _commands.GetLegendVisibility(batch, 1, shapeIndex);
        Assert.True(getHiddenResult.Success, getHiddenResult.ErrorMessage);
        Assert.False(getHiddenResult.LegendVisible);

        var showResult = _commands.SetLegendVisibility(batch, 1, shapeIndex, visible: true);
        Assert.True(showResult.Success, showResult.ErrorMessage);
        Assert.True(showResult.LegendVisible);

        var getShownResult = _commands.GetLegendVisibility(batch, 1, shapeIndex);
        Assert.True(getShownResult.Success, getShownResult.ErrorMessage);
        Assert.True(getShownResult.LegendVisible);
    }

    [Fact]
    public void SetLegendVisibility_OnShapeWithoutChart_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 50f, 50f);
            return 0;
        });

        var result = _commands.SetLegendVisibility(batch, 1, 1, visible: true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
