using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Export;
using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using Xunit.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "Composition")]
public sealed class CompositionRecipeTests(SharedPresentationFixture fixture, ITestOutputHelper output)
    : IClassFixture<SharedPresentationFixture>
{
    private readonly ShapeCommands _shapes = new();
    private readonly TextFrameCommands _text = new();

    [Theory]
    [InlineData(960)]
    [InlineData(720)]
    public void Recipes_RenderWithContentInsideSlideBounds(int slideWidth)
    {
        fixture.CreateFreshPresentation();
        var batch = fixture.Batch;
        var size = new PageSetupCommands().SetSize(batch, slideWidth, 540);
        Assert.True(size.Success, size.ErrorMessage);
        float margin = slideWidth * 0.055f;
        float contentWidth = slideWidth * 0.89f;
        float columnWidth = slideWidth * 0.42f;
        var slides = new SlideCommands();
        for (int slideIndex = 2; slideIndex <= 4; slideIndex++)
        {
            Assert.True(slides.AddBlank(batch).Success);
        }

        AddText(1, margin, 40, contentWidth, 75, "Choose the delivery model", 32, true);
        AddText(1, margin, 150, columnWidth, 50, "Shared service", 26, true);
        AddText(1, slideWidth * 0.525f, 150, columnWidth, 50, "Dedicated team", 26, true);
        AddBar(1, margin, 211, columnWidth, 4);
        AddBar(1, slideWidth * 0.525f, 211, columnWidth, 4);
        AddText(1, margin, 245, columnWidth, 175, "Start: 2 weeks\nCapacity: shared\nBest for: steady demand", 22);
        AddText(1, slideWidth * 0.525f, 245, columnWidth, 175, "Start: 6 weeks\nCapacity: reserved\nBest for: rapid iteration", 22);
        AddText(1, margin, 455, contentWidth, 45, "Decision: match the model to the workload.", 18);

        AddText(2, margin, 40, contentWidth, 75, "Adoption grows each quarter", 32, true);
        var charts = new ChartCommands();
        var chart = charts.AddChart(batch, 2, "bar", margin, 145, slideWidth * 0.57f, 285,
            ["Q1", "Q2", "Q3", "Q4"], "Active teams", [12, 18, 25, 34]);
        Assert.True(chart.Success, chart.ErrorMessage);
        Assert.True(charts.SetChartTitle(batch, 2, chart.ShapeIndex!.Value, "Active teams").Success);
        Assert.True(charts.SetLegendVisibility(batch, 2, chart.ShapeIndex.Value, false).Success);
        var data = charts.GetChartData(batch, 2, chart.ShapeIndex.Value);
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(4, data.CategoryCount);
        AddText(2, slideWidth * 0.65f, 165, slideWidth * 0.295f, 75, "+22 teams", 28, true);
        AddText(2, slideWidth * 0.65f, 255, slideWidth * 0.295f, 150,
            "Growth continues.\nNext: verify retention.", 22);
        AddText(2, margin, 455, contentWidth, 45, "Synthetic example. Quarterly active-team counts.", 16);

        AddText(3, margin, 40, contentWidth, 75, "Three steps to launch", 32, true);
        string[] periods = ["Weeks 1-2", "Weeks 3-4", "Week 5"];
        string[] milestones = ["Discover", "Pilot", "Launch"];
        string[] outcomes = ["Confirm the need", "Validate with users", "Release and monitor"];
        for (int step = 0; step < 3; step++)
        {
            float left = margin + step * slideWidth * 0.305f;
            float width = slideWidth * 0.28f;
            AddText(3, left, 160, width, 60, periods[step], 22);
            AddBar(3, left, 235, width, 5);
            AddText(3, left, 265, width, 60, milestones[step], 26, true);
            AddText(3, left, 335, width, 100, outcomes[step], 20);
        }
        AddText(3, margin, 455, contentWidth, 45, "Milestone sequence; spacing is not elapsed time.", 16);

        AddText(4, margin, 40, contentWidth, 75, "Faster time to value", 32, true);
        AddText(4, margin, 150, contentWidth, 120, "18 days", 78, true);
        AddText(4, margin, 290, contentWidth, 65, "Median time to first result", 28);
        AddBar(4, margin, 380, contentWidth, 4);
        AddText(4, margin, 410, contentWidth, 70, "Previously 24 days.\nSynthetic cohort: 40 teams, last quarter.", 20);

        string? requestedOutput = Environment.GetEnvironmentVariable("PPTMCP_RECIPE_OUTPUT");
        string directory = Path.Combine(requestedOutput ?? Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")),
            $"{slideWidth}x540");
        try
        {
            for (int slideIndex = 1; slideIndex <= 4; slideIndex++)
            {
                var notes = new NotesCommands().SetNotesText(batch, slideIndex,
                    "Synthetic composition example. All labels and data are illustrative, not customer material.");
                Assert.True(notes.Success, notes.ErrorMessage);
                string image = Path.Combine(directory, $"recipe-{slideIndex}.png");
                var exported = new ExportCommands().ExportSlideToImage(batch, slideIndex, image,
                    width: slideWidth, height: 540);
                Assert.True(exported.Success, exported.ErrorMessage);
                Assert.True(new FileInfo(image).Length > 1000);
            }
            output.WriteLine($"Recipe images: {directory}");
        }
        finally
        {
            if (requestedOutput is null && Directory.Exists(directory))
            {
                Directory.Delete(Path.GetDirectoryName(directory)!, recursive: true);
            }
        }
    }

    private void AddBar(int slideIndex, float left, float top, float width, float height)
    {
        var result = _shapes.AddRectangle(fixture.Batch, slideIndex, left, top, width, height);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(_shapes.SetFill(fixture.Batch, slideIndex, result.ShapeIndex!.Value, 0, 115, 110).Success);
        Assert.True(_shapes.SetLine(fixture.Batch, slideIndex, result.ShapeIndex.Value, visible: false).Success);
    }

    private void AddText(int slideIndex, float left, float top, float width, float height,
        string text, float fontSize, bool bold = false)
    {
        var batch = fixture.Batch;
        var result = _shapes.AddTextBox(batch, slideIndex, left, top, width, height, text);
        Assert.True(result.Success, result.ErrorMessage);
        int shapeIndex = result.ShapeIndex!.Value;
        Assert.True(_text.SetFontName(batch, slideIndex, shapeIndex, "Aptos").Success);
        Assert.True(_text.SetFontSize(batch, slideIndex, shapeIndex, fontSize).Success);
        Assert.True(_text.SetBold(batch, slideIndex, shapeIndex, bold).Success);
        Assert.True(_text.SetFontColor(batch, slideIndex, shapeIndex, 28, 35, 40).Success);
        var readBack = _text.GetText(batch, slideIndex, shapeIndex);
        Assert.True(readBack.Success, readBack.ErrorMessage);
        Assert.Equal(text, readBack.Text);
        batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            try
            {
                slides = ctx.Presentation.Slides;
                slide = slides[slideIndex];
                shapes = slide.Shapes;
                shape = shapes[shapeIndex];
                frame = shape.TextFrame;
                range = frame.TextRange;
                Assert.True(range.BoundWidth <= width + 1, $"Text too wide: {text}");
                Assert.True(range.BoundHeight <= height + 1, $"Text too tall: {text}");
                Assert.True(range.BoundLeft >= left - 1 && range.BoundTop >= top - 1);
                Assert.True(range.BoundLeft + range.BoundWidth <= left + width + 1, text);
                Assert.True(range.BoundTop + range.BoundHeight <= top + height + 1, text);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref frame);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
        });
    }
}