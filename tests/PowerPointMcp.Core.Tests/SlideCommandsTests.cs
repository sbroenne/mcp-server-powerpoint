using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Slide")]
public class SlideCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly SlideCommands _commands = new();

    [Fact]
    public void GetCount_ReturnsOne_ForFreshlyCreatedPresentation()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.GetCount(batch);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.SlideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AddBlank_IncreasesSlideCount_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var addResult = _commands.AddBlank(batch);

                Assert.True(addResult.Success);
                Assert.Null(addResult.ErrorMessage);
                Assert.Equal(2, addResult.SlideIndex);
                Assert.Equal(2, addResult.SlideCount);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var countResult = _commands.GetCount(reopened);
            Assert.Equal(2, countResult.SlideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Delete_RemovesSlide_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                _commands.AddBlank(batch); // now 2 slides
                var deleteResult = _commands.Delete(batch, 1);

                Assert.True(deleteResult.Success);
                Assert.Null(deleteResult.ErrorMessage);
                Assert.Equal(1, deleteResult.SlideCount);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var countResult = _commands.GetCount(reopened);
            Assert.Equal(1, countResult.SlideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Delete_WithInvalidIndex_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.Delete(batch, 99);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
