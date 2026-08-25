using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real PowerPoint integration tests for presentation page setup.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "PageSetup")]
public sealed class PageSetupCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PageSetupCommands _commands = new();
    private readonly PresentationCommands _presentationCommands = new();

    public PageSetupCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GetSettings_ReturnsSlideDimensionsAndFirstSlideNumber()
    {
        _fixture.CreateFreshPresentation();

        var result = _commands.GetSettings(_fixture.Batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.NotNull(result.Orientation);
        Assert.NotNull(result.FirstSlideNumber);
    }

    [Fact]
    public void SetSizeAndFirstSlideNumber_RoundTripThroughNativePageSetup()
    {
        _fixture.CreateFreshPresentation();

        var sizeResult = _commands.SetSize(_fixture.Batch, width: 720f, height: 405f);
        Assert.True(sizeResult.Success, sizeResult.ErrorMessage);
        Assert.Null(sizeResult.ErrorMessage);

        var numberResult = _commands.SetFirstSlideNumber(_fixture.Batch, firstSlideNumber: 7);
        Assert.True(numberResult.Success, numberResult.ErrorMessage);
        Assert.Null(numberResult.ErrorMessage);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var settings = _commands.GetSettings(_fixture.Batch);

        Assert.True(settings.Success, settings.ErrorMessage);
        Assert.InRange(settings.Width!.Value, 719.9f, 720.1f);
        Assert.InRange(settings.Height!.Value, 404.9f, 405.1f);
        Assert.Equal("landscape", settings.Orientation);
        Assert.Equal(7, settings.FirstSlideNumber);
    }

    [Fact]
    public void SetFooter_RoundTripsNativeHeaderFooterSettings()
    {
        _fixture.CreateFreshPresentation();

        var setResult = _commands.SetFooter(
            _fixture.Batch,
            footerText: "Release footer",
            showFooter: true,
            showSlideNumber: true,
            showDateTime: true,
            dateTimeMode: "fixed",
            fixedDateTimeText: "2026-08-25",
            showOnTitleSlide: false);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Null(setResult.ErrorMessage);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var footer = _commands.GetFooter(_fixture.Batch);

        Assert.True(footer.Success, footer.ErrorMessage);
        Assert.Equal("Release footer", footer.FooterText);
        Assert.True(footer.ShowFooter);
        Assert.True(footer.ShowSlideNumber);
        Assert.True(footer.ShowDateTime);
        Assert.Equal("fixed", footer.DateTimeMode);
        Assert.Equal("2026-08-25", footer.FixedDateTimeText);
        Assert.False(footer.ShowOnTitleSlide);
    }
}
