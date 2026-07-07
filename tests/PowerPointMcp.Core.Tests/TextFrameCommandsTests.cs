using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for text frame commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "TextFrame")]
public class TextFrameCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextFrameCommands _commands = new();

    public TextFrameCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private IPresentationBatch SetUpPresentationWithShape()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 0f, 0f, 200f, 100f);
        return batch;
    }

    [Fact]
    public void SetText_ThenGetText_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();

        var setResult = _commands.SetText(batch, 1, 1, "Quarterly Report");
        Assert.True(setResult.Success);

        var getResult = _commands.GetText(batch, 1, 1);
        Assert.True(getResult.Success);
        Assert.Equal("Quarterly Report", getResult.Text);
    }

    [Fact]
    public void SetFontSize_And_SetBold_PersistAfterSave()
    {
        var batch = SetUpPresentationWithShape();

        _commands.SetText(batch, 1, 1, "Big Bold Title");
        var sizeResult = _commands.SetFontSize(batch, 1, 1, 40f);
        Assert.True(sizeResult.Success);
        Assert.Equal(40f, sizeResult.FontSize);

        var boldResult = _commands.SetBold(batch, 1, 1, true);
        Assert.True(boldResult.Success);
        Assert.True(boldResult.Bold);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        float size = batch.Execute((ctx, ct) =>
            (float)ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Font.Size);
        Assert.Equal(40f, size);
    }

    [Fact]
    public void SetFontColor_PacksRgbInPowerPointByteOrder()
    {
        var batch = SetUpPresentationWithShape();

        _commands.SetText(batch, 1, 1, "Red Text");
        var result = _commands.SetFontColor(batch, 1, 1, red: 255, green: 0, blue: 0);

        Assert.True(result.Success);
        Assert.Equal(255, result.ColorRgb); // pure red => 0x0000FF in BGR-packed RGB
    }

    [Fact]
    public void GetText_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.GetText(batch, 1, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
