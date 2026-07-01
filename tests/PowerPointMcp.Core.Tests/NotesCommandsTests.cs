using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for speaker notes commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Notes")]
public class NotesCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly NotesCommands _commands = new();

    [Fact]
    public void SetNotesText_ThenGetNotesText_RoundTrips_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var setResult = _commands.SetNotesText(batch, 1, "Remember to mention Q3 results.");
                Assert.True(setResult.Success);
                Assert.Null(setResult.ErrorMessage);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var getResult = _commands.GetNotesText(reopened, 1);
            Assert.True(getResult.Success);
            Assert.Equal("Remember to mention Q3 results.", getResult.NotesText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetNotesText_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.GetNotesText(batch, 99);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
