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

    [Fact]
    public void SetItalic_AndGetItalic_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Italic Text");

        var setResult = _commands.SetItalic(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Italic);

        var getResult = _commands.GetItalic(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Italic);
    }

    [Fact]
    public void SetItalic_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetItalic(batch, 1, 99, true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetUnderline_AndGetUnderline_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Underlined Text");

        var setResult = _commands.SetUnderline(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Underline);

        var getResult = _commands.GetUnderline(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Underline);
    }

    [Fact]
    public void SetUnderline_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetUnderline(batch, 1, 99, true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetFontName_AndGetFontName_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Styled Text");

        var setResult = _commands.SetFontName(batch, 1, 1, "Georgia");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Georgia", setResult.FontName);

        var getResult = _commands.GetFontName(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("Georgia", getResult.FontName);
    }

    [Fact]
    public void SetFontName_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetFontName(batch, 1, 99, "Georgia");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAlignment_AndGetAlignment_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Centered Text");

        var setResult = _commands.SetAlignment(batch, 1, 1, "ppAlignCenter");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppAlignCenter", setResult.Alignment);

        var getResult = _commands.GetAlignment(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("ppAlignCenter", getResult.Alignment);
    }

    [Fact]
    public void SetAlignment_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetAlignment(batch, 1, 1, "ppAlignDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAlignment_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetAlignment(batch, 1, 99, "ppAlignCenter");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBullet_Enabled_WithCharacter_AndGetBullet_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Bulleted Text");

        var setResult = _commands.SetBullet(batch, 1, 1, enabled: true, character: "-");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.BulletEnabled);
        Assert.Equal("-", setResult.BulletCharacter);

        var getResult = _commands.GetBullet(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.BulletEnabled);
        Assert.Equal("-", getResult.BulletCharacter);
    }

    [Fact]
    public void SetBullet_Disabled_AfterEnabled_TurnsOffBullets()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        _commands.SetBullet(batch, 1, 1, enabled: true, character: "-");
        var disableResult = _commands.SetBullet(batch, 1, 1, enabled: false);
        Assert.True(disableResult.Success, disableResult.ErrorMessage);
        Assert.False(disableResult.BulletEnabled);

        var getResult = _commands.GetBullet(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.BulletEnabled);
    }

    [Fact]
    public void SetBullet_WithMultiCharacterString_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetBullet(batch, 1, 1, enabled: true, character: "ab");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBullet_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetBullet(batch, 1, 99, enabled: true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAutoSize_AndGetAutoSize_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Auto-fit Text");

        var setResult = _commands.SetAutoSize(batch, 1, 1, "ppAutoSizeShapeToFitText");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppAutoSizeShapeToFitText", setResult.AutoSize);

        var getResult = _commands.GetAutoSize(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("ppAutoSizeShapeToFitText", getResult.AutoSize);
    }

    [Fact]
    public void SetAutoSize_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetAutoSize(batch, 1, 1, "ppAutoSizeDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAutoSize_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetAutoSize(batch, 1, 99, "ppAutoSizeNone");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
