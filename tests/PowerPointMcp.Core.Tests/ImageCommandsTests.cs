using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for image commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Image")]
public class ImageCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ImageCommands _commands = new();

    public ImageCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddPicture_IncreasesShapeCount_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            var result = _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.ShapeIndex);
            Assert.Equal(1, result.ShapeCount);

            _presentationCommands.Save(batch);

            _fixture.ReopenCurrentPresentation();
            int shapeCount = batch.Execute((ctx, ct) => ctx.Presentation.Slides[1].Shapes.Count);
            Assert.Equal(1, shapeCount);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void AddPicture_WithMissingFile_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddPicture(batch, 1, "C:\\does\\not\\exist.png", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
