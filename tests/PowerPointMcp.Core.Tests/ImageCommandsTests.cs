using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for image commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Image")]
public class ImageCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ImageCommands _commands = new();

    [Fact]
    public void AddPicture_IncreasesShapeCount_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var result = _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

                Assert.True(result.Success);
                Assert.Null(result.ErrorMessage);
                Assert.Equal(1, result.ShapeIndex);
                Assert.Equal(1, result.ShapeCount);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            int shapeCount = reopened.Execute((ctx, ct) => ctx.Presentation.Slides[1].Shapes.Count);
            Assert.Equal(1, shapeCount);
        }
        finally
        {
            File.Delete(path);
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void AddPicture_WithMissingFile_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.AddPicture(batch, 1, "C:\\does\\not\\exist.png", 0f, 0f, 10f, 10f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
