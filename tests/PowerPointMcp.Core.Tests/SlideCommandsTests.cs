using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
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

    [Fact]
    public void Comments_AddListDeleteAndClear_RoundTripAgainstNativeComments()
    {
        _fixture.CreateFreshPresentation();

        var firstAdd = _commands.AddComment(
            _fixture.Batch, 1, "Release Tester", "RT", "First comment", left: 12f, top: 34f);
        Assert.True(firstAdd.Success, firstAdd.ErrorMessage);
        Assert.Equal(1, firstAdd.CommentCount);

        var secondAdd = _commands.AddComment(
            _fixture.Batch, 1, "Second Tester", "ST", "Second comment");
        Assert.True(secondAdd.Success, secondAdd.ErrorMessage);
        Assert.Equal(2, secondAdd.CommentCount);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var listed = _commands.ListComments(_fixture.Batch, 1);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(2, listed.CommentCount);
        var nativeAuthor = listed.Comments![0].Author;
        var nativeInitials = listed.Comments[0].Initials;

        // Modern PowerPoint replaces supplied comment identity with the signed-in Office identity.
        // Assert that the native identity is populated and remains stable across comments instead.
        Assert.False(string.IsNullOrWhiteSpace(nativeAuthor), "PowerPoint must expose a native comment author.");
        Assert.False(string.IsNullOrWhiteSpace(nativeInitials), "PowerPoint must expose native comment initials.");
        Assert.All(listed.Comments, comment =>
        {
            Assert.Equal(nativeAuthor, comment.Author);
            Assert.Equal(nativeInitials, comment.Initials);
        });

        Assert.Collection(
            listed.Comments,
            comment =>
            {
                Assert.Equal(1, comment.CommentIndex);
                Assert.Equal("First comment", comment.Text);
                Assert.InRange(comment.Left, 11.9f, 12.1f);
                // Modern comments may normalize the legacy COM Top value even when Left is preserved.
                Assert.True(
                    Math.Abs(comment.Top - 34f) <= 0.1f || comment.Top == 0f,
                    $"Expected the requested or native-normalized comment Top value, but found {comment.Top}.");
            },
            comment =>
            {
                Assert.Equal(2, comment.CommentIndex);
                Assert.Equal("Second comment", comment.Text);
            });

        var deleted = _commands.DeleteComment(_fixture.Batch, 1, 1);
        Assert.True(deleted.Success, deleted.ErrorMessage);
        Assert.Equal(1, deleted.CommentCount);

        saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var afterDelete = _commands.ListComments(_fixture.Batch, 1);
        var remaining = Assert.Single(afterDelete.Comments!);
        Assert.Equal("Second comment", remaining.Text);

        var cleared = _commands.ClearComments(_fixture.Batch, 1);
        Assert.True(cleared.Success, cleared.ErrorMessage);
        Assert.Equal(0, cleared.CommentCount);

        saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var afterClear = _commands.ListComments(_fixture.Batch, 1);
        Assert.True(afterClear.Success, afterClear.ErrorMessage);
        Assert.Empty(afterClear.Comments!);
    }

    [Fact]
    public void ImportFromFile_WithSourceRange_InsertsSlidesAfterDestination()
    {
        string sourcePath = _fixture.CreateFreshPresentation("pptmcp-import-source");
        _commands.AddBlank(_fixture.Batch);
        _commands.AddBlank(_fixture.Batch);
        var shapeCommands = new ShapeCommands();
        shapeCommands.AddTextBox(_fixture.Batch, 2, 10f, 10f, 200f, 40f, "Imported slide two");
        shapeCommands.AddTextBox(_fixture.Batch, 3, 10f, 10f, 200f, 40f, "Imported slide three");
        _presentationCommands.Save(_fixture.Batch);

        _fixture.CreateFreshPresentation("pptmcp-import-destination");

        var result = _commands.ImportFromFile(
            _fixture.Batch,
            sourcePath,
            destinationSlideIndex: 1,
            sourceStartSlide: 2,
            sourceEndSlide: 3);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(2, result.ImportedSlideCount);
        Assert.Equal([2, 3], result.ImportedSlideIndexes);
        Assert.Equal(3, result.SlideCount);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();
        Assert.Equal(3, _commands.GetCount(_fixture.Batch).SlideCount);

        string[] importedText = _fixture.Batch.Execute((ctx, ct) => new[]
        {
            ctx.Presentation.Slides[2].Shapes[1].TextFrame.TextRange.Text,
            ctx.Presentation.Slides[3].Shapes[1].TextFrame.TextRange.Text
        });
        Assert.Equal(["Imported slide two", "Imported slide three"], importedText);
    }

    [Fact]
    public void ImportFromFile_WithInvalidSourceRangeOrDestination_ReturnsFailure()
    {
        string sourcePath = _fixture.CreateFreshPresentation("pptmcp-import-invalid-source");
        _commands.AddBlank(_fixture.Batch);
        _presentationCommands.Save(_fixture.Batch);
        _fixture.CreateFreshPresentation("pptmcp-import-invalid-destination");

        var invalidRange = _commands.ImportFromFile(
            _fixture.Batch,
            sourcePath,
            destinationSlideIndex: 1,
            sourceStartSlide: 2,
            sourceEndSlide: 1);
        Assert.False(invalidRange.Success);
        Assert.False(string.IsNullOrWhiteSpace(invalidRange.ErrorMessage));

        var invalidDestination = _commands.ImportFromFile(
            _fixture.Batch,
            sourcePath,
            destinationSlideIndex: 99,
            sourceStartSlide: 1,
            sourceEndSlide: 1);
        Assert.False(invalidDestination.Success);
        Assert.False(string.IsNullOrWhiteSpace(invalidDestination.ErrorMessage));
    }

    [Fact]
    public void Tags_CrudIsCaseInsensitive_EnumeratesOneBased_AndPersists()
    {
        _fixture.CreateFreshPresentation();

        var firstSet = _commands.SetTag(_fixture.Batch, 1, " ReviewState ", "MiXeD Value");
        Assert.True(firstSet.Success, firstSet.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", firstSet.TagName);
        Assert.Equal("MiXeD Value", firstSet.TagValue);
        Assert.Equal(1, firstSet.TagCount);

        var updated = _commands.SetTag(_fixture.Batch, 1, " reviewstate ", "Updated Value");
        Assert.True(updated.Success, updated.ErrorMessage);
        Assert.Equal(1, updated.TagCount);

        Assert.True(_commands.SetTag(_fixture.Batch, 1, "Owner", "Alice").Success);
        Assert.True(_commands.SetTag(_fixture.Batch, 1, "ReviewState", "Unspaced Value").Success);

        var get = _commands.GetTag(_fixture.Batch, 1, " ReViEwStAtE ");
        Assert.True(get.Success, get.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", get.TagName);
        Assert.Equal("Updated Value", get.TagValue);
        Assert.Equal(1, get.TagIndex);

        var unspacedGet = _commands.GetTag(_fixture.Batch, 1, "reviewstate");
        Assert.True(unspacedGet.Success, unspacedGet.ErrorMessage);
        Assert.Equal("Unspaced Value", unspacedGet.TagValue);
        Assert.Equal(3, unspacedGet.TagIndex);

        var listed = _commands.ListTags(_fixture.Batch, 1);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(3, listed.TagCount);
        Assert.Equal([1, 2, 3], listed.Tags!.Select(tag => tag.TagIndex));
        Assert.Equal([" REVIEWSTATE ", "OWNER", "REVIEWSTATE"], listed.Tags!.Select(tag => tag.Name));

        Assert.True(_presentationCommands.Save(_fixture.Batch).Success);
        _fixture.ReopenCurrentPresentation();
        Assert.Equal("Updated Value", _commands.GetTag(_fixture.Batch, 1, " reviewstate ").TagValue);

        var deleted = _commands.DeleteTag(_fixture.Batch, 1, " REVIEWSTATE ");
        Assert.True(deleted.Success, deleted.ErrorMessage);
        Assert.Equal(2, deleted.TagCount);
        Assert.False(_commands.GetTag(_fixture.Batch, 1, " reviewstate ").Success);
        Assert.False(_commands.DeleteTag(_fixture.Batch, 1, " reviewstate ").Success);
        Assert.True(_commands.GetTag(_fixture.Batch, 1, "reviewstate").Success);

        Assert.True(_commands.ListTags(_fixture.Batch, 1).Success);
        Assert.Equal(1, _commands.GetCount(_fixture.Batch).SlideCount);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(99)]
    public void Tags_WithInvalidSlideIndex_ReturnFailure(int slideIndex)
    {
        _fixture.CreateFreshPresentation();

        Assert.False(_commands.SetTag(_fixture.Batch, slideIndex, "name", "value").Success);
        Assert.False(_commands.GetTag(_fixture.Batch, slideIndex, "name").Success);
        Assert.False(_commands.ListTags(_fixture.Batch, slideIndex).Success);
        Assert.False(_commands.DeleteTag(_fixture.Batch, slideIndex, "name").Success);
    }
}
