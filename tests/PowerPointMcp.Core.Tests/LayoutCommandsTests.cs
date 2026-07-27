using Sbroenne.PowerPointMcp.Core.Layout;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide layout commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Layout")]
public class LayoutCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly LayoutCommands _commands = new();

    public LayoutCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetLayout_ThenGetLayout_RoundTrips_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetLayout(batch, 1, "ppLayoutTitleOnly");
        Assert.True(setResult.Success);
        Assert.Equal("ppLayoutTitleOnly", setResult.LayoutName);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var getResult = _commands.GetLayout(batch, 1);
        Assert.True(getResult.Success);
        Assert.Equal("ppLayoutTitleOnly", getResult.LayoutName);
    }

    [Fact]
    public void SetLayout_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetLayout(batch, 1, "NotARealLayout");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ListLayouts_ReturnsLayoutsWithUsageFlag()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.ListLayouts(batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.MasterIndex);
        Assert.NotNull(result.Layouts);
        Assert.NotEmpty(result.Layouts);
        Assert.Contains(result.Layouts!, layout => layout.IsUsed);
    }

    [Fact]
    public void DeleteLayout_RemovesUnusedLayout()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            PowerPoint.Design design = ctx.Presentation.Designs[1];
            PowerPoint.Master slideMaster = design.SlideMaster;
            PowerPoint.CustomLayouts customLayouts = slideMaster.CustomLayouts;
            PowerPoint.CustomLayout sourceLayout = customLayouts[1];
            PowerPoint.CustomLayout duplicateLayout = sourceLayout.Duplicate();
            duplicateLayout.Name = "PptMcpTestUnusedLayout";
            return 0;
        });

        var listBefore = _commands.ListLayouts(batch, 1);
        Assert.True(listBefore.Success, listBefore.ErrorMessage);
        var candidateLayout = listBefore.Layouts!.Single(layout => layout.LayoutName == "PptMcpTestUnusedLayout");

        var deleteResult = _commands.DeleteLayout(batch, 1, candidateLayout.LayoutIndex);
        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);

        var listAfter = _commands.ListLayouts(batch, 1);
        Assert.True(listAfter.Success, listAfter.ErrorMessage);
        Assert.DoesNotContain(listAfter.Layouts!, layout => layout.LayoutName == candidateLayout.LayoutName);
        Assert.Equal(listBefore.Layouts!.Count - 1, listAfter.Layouts!.Count);
    }
}
