using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Layout;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide layout commands against live PowerPoint COM. No mocking.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Layout")]
public class LayoutCommandsTests
{
    private readonly PresentationCommands _presentationCommands = new();
    private readonly LayoutCommands _commands = new();

    [Fact]
    public void SetLayout_ThenGetLayout_RoundTrips_AndPersistsAfterSave()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                var setResult = _commands.SetLayout(batch, 1, "ppLayoutTitleOnly");
                Assert.True(setResult.Success);
                Assert.Equal("ppLayoutTitleOnly", setResult.LayoutName);

                _presentationCommands.Save(batch);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            var getResult = _commands.GetLayout(reopened, 1);
            Assert.True(getResult.Success);
            Assert.Equal("ppLayoutTitleOnly", getResult.LayoutName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetLayout_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _presentationCommands.Create(path);

            using var batch = PresentationSession.BeginBatch(path);
            var result = _commands.SetLayout(batch, 1, "NotARealLayout");

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
