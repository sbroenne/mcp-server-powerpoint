using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for speaker notes commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Notes")]
public class NotesCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly NotesCommands _commands = new();

    public NotesCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetNotesText_ThenGetNotesText_RoundTrips_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetNotesText(batch, 1, "Remember to mention Q3 results.");
        Assert.True(setResult.Success);
        Assert.Null(setResult.ErrorMessage);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var getResult = _commands.GetNotesText(batch, 1);
        Assert.True(getResult.Success);
        Assert.Equal("Remember to mention Q3 results.", getResult.NotesText);
    }

    [Fact]
    public void GetNotesText_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetNotesText(batch, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
