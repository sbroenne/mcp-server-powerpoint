using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for restyling an existing presentation via a PowerPoint .potx
/// template (<see cref="PresentationCommands.ApplyTemplate"/> /
/// <see cref="PresentationCommands.GetThemeName"/>). No mocking — drives live PowerPoint COM.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/>; the .potx template asset itself is built once, at
/// test-assembly load time, in full isolation — see <see cref="SharedTemplateAsset"/> for why
/// that isolation matters (a real, previously-hung-then-mis-diagnosed COM process-sharing bug).
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Presentation")]
public class ApplyTemplateTests : IClassFixture<SharedPresentationFixture>
{
    private const string TemplateDesignName = SharedTemplateAsset.DesignName;

    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _commands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextFrameCommands _textFrameCommands = new();

    public ApplyTemplateTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ApplyTemplate_ToDeckWithContent_PreservesSlideContent_AndUpdatesThemeName()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        const string knownText = "Do not lose me when the template changes.";
        _shapeCommands.AddRectangle(batch, 1, 0f, 0f, 200f, 100f);
        _textFrameCommands.SetText(batch, 1, 1, knownText);

        var applyResult = _commands.ApplyTemplate(batch, SharedTemplateAsset.TemplatePath);

        Assert.True(applyResult.Success, applyResult.ErrorMessage);
        Assert.Null(applyResult.ErrorMessage);
        Assert.Equal(TemplateDesignName, applyResult.ThemeName);

        // Slide content must survive the restyle.
        var textAfter = _textFrameCommands.GetText(batch, 1, 1);
        Assert.True(textAfter.Success);
        Assert.Equal(knownText, textAfter.Text);

        var themeResult = _commands.GetThemeName(batch);
        Assert.True(themeResult.Success);
        Assert.Equal(TemplateDesignName, themeResult.ThemeName);
    }

    [Fact]
    public void ApplyTemplate_WithMissingFile_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        string missingPath = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"does-not-exist-{Guid.NewGuid():N}.potx");

        var result = _commands.ApplyTemplate(batch, missingPath);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ApplyTemplate_WithUnsupportedExtension_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        string unsupportedPath = CoreTestHelper.CreateUniqueTestImageFile();

        var result = _commands.ApplyTemplate(batch, unsupportedPath);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Contains(".png", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ApplyTemplate_WithEmptyPath_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.ApplyTemplate(batch, string.Empty);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetThemeName_OnFreshPresentation_ReturnsSuccessWithNonEmptyName()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetThemeName(batch);

        Assert.True(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ThemeName));
    }
}
