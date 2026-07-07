using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Slide")]
public class SlideCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly SlideCommands _commands = new();

    public SlideCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GetCount_ReturnsOne_ForFreshlyCreatedPresentation()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetCount(batch);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.SlideCount);
    }

    [Fact]
    public void AddBlank_IncreasesSlideCount_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var addResult = _commands.AddBlank(batch);

        Assert.True(addResult.Success);
        Assert.Null(addResult.ErrorMessage);
        Assert.Equal(2, addResult.SlideIndex);
        Assert.Equal(2, addResult.SlideCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch);
        Assert.Equal(2, countResult.SlideCount);
    }

    [Fact]
    public void Delete_RemovesSlide_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddBlank(batch); // now 2 slides
        var deleteResult = _commands.Delete(batch, 1);

        Assert.True(deleteResult.Success);
        Assert.Null(deleteResult.ErrorMessage);
        Assert.Equal(1, deleteResult.SlideCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch);
        Assert.Equal(1, countResult.SlideCount);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.Delete(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
