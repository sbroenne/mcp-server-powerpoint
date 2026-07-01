using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for shape commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Shape")]
public class ShapeCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _commands = new();

    [Fact]
    public void AddRectangle_IncreasesShapeCount_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var result = _commands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

                Assert.True(result.Success);
                Assert.Null(result.ErrorMessage);
                Assert.Equal(1, result.ShapeIndex);
                Assert.Equal(1, result.ShapeCount);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var countResult = _commands.GetCount(reopened, 1);
            Assert.Equal(1, countResult.ShapeCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AddTextBox_SetsText_ReadableAfterReopen()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var result = _commands.AddTextBox(batch, 1, 0f, 0f, 200f, 40f, "Hello PowerPoint");
                Assert.True(result.Success);
                Assert.Equal(1, result.ShapeIndex);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            string text = reopened.Execute((ctx, ct) =>
                ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Text);
            Assert.Equal("Hello PowerPoint", text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetPositionAndSize_UpdatesShapeGeometry()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

            var posResult = _commands.SetPosition(batch, 1, 1, 123f, 456f);
            Assert.True(posResult.Success);
            Assert.Equal(123f, posResult.Left);
            Assert.Equal(456f, posResult.Top);

            var sizeResult = _commands.SetSize(batch, 1, 1, 300f, 200f);
            Assert.True(sizeResult.Success);
            Assert.Equal(300f, sizeResult.Width);
            Assert.Equal(200f, sizeResult.Height);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Delete_RemovesShape_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
                var deleteResult = _commands.Delete(batch, 1, 1);

                Assert.True(deleteResult.Success);
                Assert.Equal(0, deleteResult.ShapeCount);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var countResult = _commands.GetCount(reopened, 1);
            Assert.Equal(0, countResult.ShapeCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AddRectangle_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.AddRectangle(batch, 99, 0f, 0f, 10f, 10f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
