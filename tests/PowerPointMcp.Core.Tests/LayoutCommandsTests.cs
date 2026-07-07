using Sbroenne.PowerPointMcp.Core.Layout;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide layout commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Layout")]
public class LayoutCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly LayoutCommands _commands = new();

    public LayoutCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetLayout_ThenGetLayout_RoundTrips_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetLayout(batch, 1, "ppLayoutTitleOnly");
        Assert.True(setResult.Success);
        Assert.Equal("ppLayoutTitleOnly", setResult.LayoutName);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var getResult = _commands.GetLayout(batch, 1);
        Assert.True(getResult.Success);
        Assert.Equal("ppLayoutTitleOnly", getResult.LayoutName);
    }

    [Fact]
    public void SetLayout_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetLayout(batch, 1, "NotARealLayout");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
