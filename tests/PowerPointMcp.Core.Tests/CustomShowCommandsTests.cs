using Sbroenne.PowerPointMcp.Core.CustomShow;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real PowerPoint integration tests for named custom slide shows.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CustomShow")]
public sealed class CustomShowCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly CustomShowCommands _commands = new();
    private readonly SlideCommands _slideCommands = new();
    private readonly PresentationCommands _presentationCommands = new();

    public CustomShowCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    /// <summary>Adds blank slides until the presentation has exactly <paramref name="count"/> slides.</summary>
    private int EnsureSlideCount(int count)
    {
        var countResult = _slideCommands.GetCount(_fixture.Batch);
        Assert.True(countResult.Success, countResult.ErrorMessage);
        int current = countResult.SlideCount!.Value;
        while (current < count)
        {
            var result = _slideCommands.AddBlank(_fixture.Batch);
            Assert.True(result.Success, result.ErrorMessage);
            current = result.SlideCount!.Value;
        }

        return current;
    }

    [Fact]
    public void List_ReturnsEmptyCollection_WhenNoCustomShowsExist()
    {
        _fixture.CreateFreshPresentation();

        var result = _commands.List(_fixture.Batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.Shows);
        Assert.Empty(result.Shows!);
    }

    [Fact]
    public void Create_AddsNamedShow_AndListReflectsItInOrder()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(4);

        var createResult = _commands.Create(_fixture.Batch, "Executive Summary", [3, 1]);

        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.Null(createResult.ErrorMessage);
        Assert.Equal("Executive Summary", createResult.Name);
        Assert.Equal([3, 1], createResult.SlideIndices);

        var listResult = _commands.List(_fixture.Batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        var entry = Assert.Single(listResult.Shows!);
        Assert.Equal(1, entry.Index);
        Assert.Equal("Executive Summary", entry.Name);
        Assert.Equal([3, 1], entry.SlideIndices);
    }

    [Fact]
    public void List_ReturnsMultipleShows_InCollectionOrderWithDistinctSlideIndices()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(4);

        var first = _commands.Create(_fixture.Batch, "Executive Summary", [1, 2]);
        Assert.True(first.Success, first.ErrorMessage);
        var second = _commands.Create(_fixture.Batch, "Technical Deep Dive", [3, 4, 1]);
        Assert.True(second.Success, second.ErrorMessage);

        var listResult = _commands.List(_fixture.Batch);

        Assert.True(listResult.Success, listResult.ErrorMessage);
        Assert.Equal(2, listResult.Shows!.Count);
        Assert.Equal(1, listResult.Shows![0].Index);
        Assert.Equal("Executive Summary", listResult.Shows![0].Name);
        Assert.Equal([1, 2], listResult.Shows![0].SlideIndices);
        Assert.Equal(2, listResult.Shows![1].Index);
        Assert.Equal("Technical Deep Dive", listResult.Shows![1].Name);
        Assert.Equal([3, 4, 1], listResult.Shows![1].SlideIndices);
    }

    [Fact]
    public void Create_AllowsRepeatedSlideIndex()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);

        var createResult = _commands.Create(_fixture.Batch, "Repeats Slide One", [1, 2, 1]);

        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.Equal([1, 2, 1], createResult.SlideIndices);

        var listResult = _commands.List(_fixture.Batch);
        var entry = Assert.Single(listResult.Shows!);
        Assert.Equal([1, 2, 1], entry.SlideIndices);
    }

    [Fact]
    public void Create_WithDuplicateName_ReturnsFailureWithoutAddingSecondShow()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);

        var first = _commands.Create(_fixture.Batch, "Demo Flow", [1]);
        Assert.True(first.Success, first.ErrorMessage);

        var second = _commands.Create(_fixture.Batch, "Demo Flow", [2]);

        Assert.False(second.Success);
        Assert.Contains("already exists", second.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Single(listResult.Shows!);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_WithWhitespaceName_ReturnsFailureWithoutAddingShow(string invalidName)
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(1);

        var result = _commands.Create(_fixture.Batch, invalidName, [1]);

        Assert.False(result.Success);
        Assert.Contains("name is required", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Empty(listResult.Shows!);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(5)]
    public void Create_WithOutOfRangeSlideIndex_ReturnsFailureWithoutAddingShow(int invalidIndex)
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);

        var result = _commands.Create(_fixture.Batch, "Bad Range", [invalidIndex]);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Empty(listResult.Shows!);
    }

    [Fact]
    public void Create_WithEmptySlideIndices_ReturnsFailureWithoutAddingShow()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(1);

        var result = _commands.Create(_fixture.Batch, "Empty Show", Array.Empty<int>());

        Assert.False(result.Success);
        Assert.Contains("at least one", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Empty(listResult.Shows!);
    }

    [Fact]
    public void Delete_RemovesNamedShow_AndListNoLongerIncludesIt()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);
        var createResult = _commands.Create(_fixture.Batch, "Temporary Show", [1, 2]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var deleteResult = _commands.Delete(_fixture.Batch, "Temporary Show");

        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);
        Assert.Null(deleteResult.ErrorMessage);
        Assert.Equal("Temporary Show", deleteResult.Name);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Empty(listResult.Shows!);
    }

    [Fact]
    public void Delete_WithUnknownName_ReturnsFailureWithoutMutatingExistingShows()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);
        var createResult = _commands.Create(_fixture.Batch, "Keep Me", [1]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var deleteResult = _commands.Delete(_fixture.Batch, "Does Not Exist");

        Assert.False(deleteResult.Success);
        Assert.Contains("No custom show", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        var entry = Assert.Single(listResult.Shows!);
        Assert.Equal("Keep Me", entry.Name);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Delete_WithWhitespaceName_ReturnsFailureWithoutMutatingExistingShows(string invalidName)
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);
        var createResult = _commands.Create(_fixture.Batch, "Keep Me Too", [1]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var deleteResult = _commands.Delete(_fixture.Batch, invalidName);

        Assert.False(deleteResult.Success);
        Assert.Contains("name is required", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Single(listResult.Shows!);
    }

    [Fact]
    public void Create_PersistsAcrossSaveAndReopen()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(3);
        var createResult = _commands.Create(_fixture.Batch, "Persisted Show", [2, 3, 1]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var listResult = _commands.List(_fixture.Batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        var entry = Assert.Single(listResult.Shows!);
        Assert.Equal("Persisted Show", entry.Name);
        Assert.Equal([2, 3, 1], entry.SlideIndices);
    }

    [Fact]
    public void List_ResolvesCurrentSlideIndices_AfterReorderingSlides()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(3);
        var createResult = _commands.Create(_fixture.Batch, "Follows Slide Identity", [1, 3]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        // Move slide 1 to the end: original slide 1 is now at position 3.
        var moveResult = _slideCommands.MoveTo(_fixture.Batch, slideIndex: 1, toPosition: 3);
        Assert.True(moveResult.Success, moveResult.ErrorMessage);

        var listResult = _commands.List(_fixture.Batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        var entry = Assert.Single(listResult.Shows!);
        // The show still references the same two slides by identity, now at their new positions.
        Assert.Equal([3, 2], entry.SlideIndices);
    }

    [Fact]
    public void List_OmitsStaleSlideId_WhenAReferencedSlideIsLaterDeleted()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(3);
        var createResult = _commands.Create(_fixture.Batch, "Survives Slide Deletion", [1, 2, 3]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var deleteResult = _slideCommands.Delete(_fixture.Batch, 2);
        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);

        var listResult = _commands.List(_fixture.Batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        var entry = Assert.Single(listResult.Shows!);
        // The deleted slide's stale ID is omitted; the other two (now renumbered) remain.
        Assert.Equal([1, 2], entry.SlideIndices);
    }

    [Fact]
    public void Create_WithNameDifferingOnlyByCase_ReturnsFailureWithoutAddingSecondShow()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(2);

        var first = _commands.Create(_fixture.Batch, "Demo Flow", [1]);
        Assert.True(first.Success, first.ErrorMessage);

        var second = _commands.Create(_fixture.Batch, "DEMO FLOW", [2]);

        Assert.False(second.Success);
        Assert.Contains("already exists", second.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Single(listResult.Shows!);
    }

    [Fact]
    public void Delete_WithNameDifferingOnlyByCase_RemovesNamedShow()
    {
        _fixture.CreateFreshPresentation();
        EnsureSlideCount(1);
        var createResult = _commands.Create(_fixture.Batch, "Demo Flow", [1]);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var deleteResult = _commands.Delete(_fixture.Batch, "DEMO FLOW");

        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);

        var listResult = _commands.List(_fixture.Batch);
        Assert.Empty(listResult.Shows!);
    }
}
