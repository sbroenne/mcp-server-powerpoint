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

    [Fact]
    public void Duplicate_InsertsCopyImmediatelyAfterSource_AndIncreasesCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddBlank(batch); // now 2 slides

        var result = _commands.Duplicate(batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.SlideIndex);
        Assert.Equal(3, result.SlideCount);
    }

    [Fact]
    public void Duplicate_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.Duplicate(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void MoveTo_ReordersSlide_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddBlank(batch);
        _commands.AddBlank(batch); // now 3 slides

        var moveResult = _commands.MoveTo(batch, 1, 3);

        Assert.True(moveResult.Success, moveResult.ErrorMessage);
        Assert.Equal(3, moveResult.SlideIndex);
        Assert.Equal(3, moveResult.SlideCount);

        _presentationCommands.Save(batch);
        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch);
        Assert.Equal(3, countResult.SlideCount);
    }

    [Fact]
    public void MoveTo_WithInvalidToPosition_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddBlank(batch); // now 2 slides

        var result = _commands.MoveTo(batch, 1, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBackgroundColor_AndGetBackgroundColor_RoundTripsColorAndFollowsMasterFlag()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var beforeResult = _commands.GetBackgroundColor(batch, 1);
        Assert.True(beforeResult.Success, beforeResult.ErrorMessage);
        Assert.True(beforeResult.FollowsMasterBackground);

        var setResult = _commands.SetBackgroundColor(batch, 1, red: 0, green: 0, blue: 255);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(16711680, setResult.ColorRgb);
        Assert.False(setResult.FollowsMasterBackground);

        var getResult = _commands.GetBackgroundColor(batch, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(16711680, getResult.ColorRgb);
        Assert.False(getResult.FollowsMasterBackground);
    }

    [Fact]
    public void SetBackgroundColor_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetBackgroundColor(batch, 99, red: 255, green: 0, blue: 0);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetGradientBackground_AndGetGradientBackground_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetGradientBackground(
            batch, 1,
            red1: 255, green1: 0, blue1: 0,
            red2: 0, green2: 0, blue2: 255,
            gradientStyle: "msoGradientVertical",
            gradientVariant: 2);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);
        Assert.Equal(16711680, setResult.ColorRgb2);
        Assert.Equal("msoGradientVertical", setResult.GradientStyleName);
        Assert.Equal(2, setResult.GradientVariant);
        Assert.False(setResult.FollowsMasterBackground);

        var getResult = _commands.GetGradientBackground(batch, 1);
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
            batch, 1,
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

        _commands.SetBackgroundColor(batch, 1, red: 255, green: 0, blue: 0);

        var result = _commands.GetGradientBackground(batch, 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddSection_AppendsSection_AndIncreasesSectionCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddSection(batch, 1, "Introduction");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.SectionIndex);
        Assert.Equal(1, result.SectionCount);
    }

    [Fact]
    public void AddSection_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddSection(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void RenameSection_ThenGetSectionName_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddSection(batch, 1, "Original");

        var renameResult = _commands.RenameSection(batch, 1, "Renamed");
        Assert.True(renameResult.Success, renameResult.ErrorMessage);

        var getNameResult = _commands.GetSectionName(batch, 1);
        Assert.True(getNameResult.Success, getNameResult.ErrorMessage);
        Assert.Equal("Renamed", getNameResult.SectionName);
    }

    [Fact]
    public void RenameSection_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.RenameSection(batch, 99, "x");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetSectionCount_ReturnsZero_WhenNoSectionsExist()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetSectionCount(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(0, result.SectionCount);
    }

    [Fact]
    public void DeleteSection_KeepingSlides_DecreasesSectionCount_ButKeepsSlideCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddBlank(batch); // 2 slides
        _commands.AddSection(batch, 1, "Section A");
        _commands.AddSection(batch, 2, "Section B");

        // PowerPoint disallows deleting section 1 unless deleteSlides is true, so delete
        // section 2 here to exercise the keep-slides path.
        var result = _commands.DeleteSection(batch, 2, deleteSlides: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.SectionCount);
        Assert.Equal(2, _commands.GetCount(batch).SlideCount);
    }

    [Fact]
    public void DeleteSection_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.DeleteSection(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetSectionName_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetSectionName(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
