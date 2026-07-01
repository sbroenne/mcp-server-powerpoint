using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for text frame commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "TextFrame")]
public class TextFrameCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextFrameCommands _commands = new();

    private string SetUpPresentationWithShape()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        _presentationCommands.Create(path);
        using var batch = PresentationSession.BeginBatch(path);
        _shapeCommands.AddRectangle(batch, 1, 0f, 0f, 200f, 100f);
        _presentationCommands.Save(batch);
        return path;
    }

    [Fact]
    public void SetText_ThenGetText_RoundTrips()
    {
        string path = SetUpPresentationWithShape();
        try
        {
            using var batch = PresentationSession.BeginBatch(path);
            var setResult = _commands.SetText(batch, 1, 1, "Quarterly Report");
            Assert.True(setResult.Success);

            var getResult = _commands.GetText(batch, 1, 1);
            Assert.True(getResult.Success);
            Assert.Equal("Quarterly Report", getResult.Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetFontSize_And_SetBold_PersistAfterSave()
    {
        string path = SetUpPresentationWithShape();
        try
        {
            using (var batch = PresentationSession.BeginBatch(path))
            {
                _commands.SetText(batch, 1, 1, "Big Bold Title");
                var sizeResult = _commands.SetFontSize(batch, 1, 1, 40f);
                Assert.True(sizeResult.Success);
                Assert.Equal(40f, sizeResult.FontSize);

                var boldResult = _commands.SetBold(batch, 1, 1, true);
                Assert.True(boldResult.Success);
                Assert.True(boldResult.Bold);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            float size = reopened.Execute((ctx, ct) =>
                (float)ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Font.Size);
            Assert.Equal(40f, size);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetFontColor_PacksRgbInPowerPointByteOrder()
    {
        string path = SetUpPresentationWithShape();
        try
        {
            using var batch = PresentationSession.BeginBatch(path);
            _commands.SetText(batch, 1, 1, "Red Text");
            var result = _commands.SetFontColor(batch, 1, 1, red: 255, green: 0, blue: 0);

            Assert.True(result.Success);
            Assert.Equal(255, result.ColorRgb); // pure red => 0x0000FF in BGR-packed RGB
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetText_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        string path = SetUpPresentationWithShape();
        try
        {
            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.GetText(batch, 1, 99);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
