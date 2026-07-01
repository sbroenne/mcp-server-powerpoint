using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests against a live PowerPoint COM instance. NO mocking — per Rule 30
/// (integration tests over unit tests), these require PowerPoint to be installed and drive
/// the actual Presentations.Add/Open/SaveAs/Close/Quit COM calls.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Presentation")]
public class PresentationCommandsTests
{
    private readonly PresentationCommands _commands = new();

    [Fact]
    public void Create_SavesRealPptxFile_ThatPowerPointCanReopen()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            var result = _commands.Create(path);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.True(File.Exists(path), "Create() must produce a real .pptx file on disk.");
            Assert.True(new FileInfo(path).Length > 0, "The saved .pptx must not be empty.");

            // Round-trip: open the file we just created with a fresh PowerPoint COM session
            // and verify it is a valid, readable presentation with the default single slide.
            using var batch = PresentationSession.BeginBatch(path);
            int slideCount = batch.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
            Assert.Equal(1, slideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Save_PersistsChanges_VisibleAfterReopen()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                batch.Execute((ctx, ct) =>
                {
                    ctx.Presentation.Slides.Add(2, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
                    return 0;
                });

                var saveResult = _commands.Save(batch);
                Assert.True(saveResult.Success);
                Assert.Null(saveResult.ErrorMessage);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            int slideCount = reopened.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
            Assert.Equal(2, slideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
