using Sbroenne.PowerPointMcp.Core.Master;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide master title/body font and background color operations
/// (<see cref="MasterCommands"/>). No mocking — drives live PowerPoint COM. Shares one
/// PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/>.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Master")]
public class MasterCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly MasterCommands _commands = new();

    public MasterCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GetTitleFont_OnFreshPresentation_ReturnsSuccessWithFontDetails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetTitleFont(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.False(string.IsNullOrEmpty(result.FontName));
        Assert.NotNull(result.FontSize);
        Assert.NotNull(result.Bold);
        Assert.NotNull(result.ColorRgb);
    }

    [Fact]
    public void SetTitleFont_ChangesNameSizeBoldAndColor_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetTitleFont(batch, fontName: "Arial", fontSize: 44f, bold: true, red: 200, green: 30, blue: 30);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Arial", setResult.FontName);
        Assert.Equal(44f, setResult.FontSize);
        Assert.True(setResult.Bold);

        var getResult = _commands.GetTitleFont(batch);
        Assert.True(getResult.Success);
        Assert.Equal("Arial", getResult.FontName);
        Assert.Equal(44f, getResult.FontSize);
        Assert.True(getResult.Bold);
        int expectedRgb = 200 + (30 << 8) + (30 << 16);
        Assert.Equal(expectedRgb, getResult.ColorRgb);
    }

    [Fact]
    public void SetTitleFont_WithOnlyFontSize_LeavesOtherPropertiesUnchanged()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var before = _commands.GetTitleFont(batch);
        Assert.True(before.Success);

        var setResult = _commands.SetTitleFont(batch, fontSize: 60f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(60f, setResult.FontSize);
        Assert.Equal(before.FontName, setResult.FontName);
        Assert.Equal(before.Bold, setResult.Bold);
    }

    [Fact]
    public void GetBodyFont_OnFreshPresentation_ReturnsSuccessWithFontDetails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetBodyFont(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrEmpty(result.FontName));
    }

    [Fact]
    public void SetBodyFont_ChangesNameAndSize_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetBodyFont(batch, fontName: "Georgia", fontSize: 22f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Georgia", setResult.FontName);
        Assert.Equal(22f, setResult.FontSize);

        var getResult = _commands.GetBodyFont(batch);
        Assert.True(getResult.Success);
        Assert.Equal("Georgia", getResult.FontName);
        Assert.Equal(22f, getResult.FontSize);
    }

    [Fact]
    public void SetAndGetBackgroundColor_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetBackgroundColor(batch, red: 10, green: 20, blue: 30);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        int expectedRgb = 10 + (20 << 8) + (30 << 16);
        Assert.Equal(expectedRgb, setResult.ColorRgb);

        var getResult = _commands.GetBackgroundColor(batch);
        Assert.True(getResult.Success);
        Assert.Equal(expectedRgb, getResult.ColorRgb);
    }

    [Fact]
    public void SetGradientBackground_AndGetGradientBackground_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetGradientBackground(
            batch,
            red1: 255, green1: 0, blue1: 0,
            red2: 0, green2: 0, blue2: 255,
            gradientStyle: "msoGradientVertical",
            gradientVariant: 2);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);
        Assert.Equal(16711680, setResult.ColorRgb2);
        Assert.Equal("msoGradientVertical", setResult.GradientStyleName);
        Assert.Equal(2, setResult.GradientVariant);

        var getResult = _commands.GetGradientBackground(batch);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(255, getResult.ColorRgb);
        Assert.Equal(16711680, getResult.ColorRgb2);
        Assert.Equal("msoGradientVertical", getResult.GradientStyleName);
        Assert.Equal(2, getResult.GradientVariant);
    }

    [Fact]
    public void SetGradientBackground_WithUnrecognizedStyleName_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetGradientBackground(
            batch,
            red1: 255, green1: 0, blue1: 0,
            red2: 0, green2: 0, blue2: 255,
            gradientStyle: "msoGradientNotARealStyle");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetGradientBackground_WhenBackgroundIsSolid_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.SetBackgroundColor(batch, red: 255, green: 0, blue: 0);

        var result = _commands.GetGradientBackground(batch);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
