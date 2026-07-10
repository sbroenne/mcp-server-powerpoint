using Sbroenne.PowerPointMcp.Core.SmartArt;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for SmartArt diagram creation and node manipulation
/// (<see cref="SmartArtCommands"/>). No mocking — drives live PowerPoint COM. Shares one
/// PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/>.
/// </summary>
/// <remarks>
/// Unlike most other Core test classes, these tests deliberately do NOT call
/// <see cref="SharedPresentationFixture.CreateFreshPresentation"/> between [Fact]s. Verified
/// live: closing/reopening the presentation (<c>PresentationBatch.ReopenPresentation</c>) followed
/// by an <c>Application.SmartArtLayouts</c> access reliably throws
/// <see cref="UnauthorizedAccessException"/> (E_ACCESSDENIED) — a real Office COM quirk specific
/// to the SmartArt gallery, not a transient/retryable condition (confirmed via a 15s retry budget
/// that did not resolve it). Isolation is achieved instead by adding a fresh blank slide per test
/// and placing that test's SmartArt shape(s) only on that slide — the same live Application/
/// presentation is reused throughout, and <c>SmartArtLayouts</c> is only ever queried against the
/// ORIGINAL presentation instance, which does not exhibit the bug.
/// </remarks>
[Trait("Category", "Integration")]
[Trait("Feature", "SmartArt")]
public class SmartArtCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly SmartArtCommands _commands = new();

    public SmartArtCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    /// <summary>Adds a fresh blank slide to the shared presentation and returns its 1-based index.</summary>
    private int AddIsolatedSlide()
    {
        var batch = _fixture.Batch;
        return batch.Execute((ctx, ct) =>
        {
            int newIndex = ctx.Presentation.Slides.Count + 1;
            ctx.Presentation.Slides.Add(newIndex, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
            return newIndex;
        });
    }

    [Fact]
    public void AddSmartArt_WithBasicProcessLayout_ReturnsSuccessWithNodes()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var result = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 50, top: 50, width: 500, height: 300);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal("Basic Process", result.LayoutName);
        Assert.NotNull(result.NodeCount);
        Assert.True(result.NodeCount > 0);
    }

    [Fact]
    public void AddSmartArt_WithUnknownLayoutName_ReturnsFailure()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var result = _commands.AddSmartArt(batch, slideIndex, layoutName: "Not A Real Layout", left: 50, top: 50, width: 500, height: 300);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void AddSmartArt_WithInvalidSlideIndex_ReturnsFailure()
    {
        var batch = _fixture.Batch;
        AddIsolatedSlide();

        var result = _commands.AddSmartArt(batch, slideIndex: 9999, layoutName: "Basic Process", left: 0, top: 0, width: 100, height: 100);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void GetNodeCount_OnFreshBasicProcessDiagram_MatchesDefaultNodes()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 0, top: 0, width: 400, height: 200);
        Assert.True(addResult.Success, addResult.ErrorMessage);

        var countResult = _commands.GetNodeCount(batch, slideIndex, shapeIndex: addResult.ShapeIndex!.Value);

        Assert.True(countResult.Success, countResult.ErrorMessage);
        Assert.Equal(addResult.NodeCount, countResult.NodeCount);
    }

    [Fact]
    public void GetNodeCount_OnShapeWithoutSmartArt_ReturnsFailure()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        batch.Execute((ctx, ct) =>
        {
            dynamic s = ctx.Presentation.Slides[slideIndex];
            s.Shapes.AddShape(1 /* msoShapeRectangle */, 0f, 0f, 100f, 100f);
        });

        var result = _commands.GetNodeCount(batch, slideIndex, shapeIndex: 1);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void SetNodeText_ThenGetNodeText_RoundTrips()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 0, top: 0, width: 400, height: 200);
        Assert.True(addResult.Success, addResult.ErrorMessage);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var setResult = _commands.SetNodeText(batch, slideIndex, shapeIndex, nodeIndex: 1, text: "Kick off");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Kick off", setResult.NodeText);

        var getResult = _commands.GetNodeText(batch, slideIndex, shapeIndex, nodeIndex: 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("Kick off", getResult.NodeText);
    }

    [Fact]
    public void SetNodeText_WithInvalidNodeIndex_ReturnsFailure()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 0, top: 0, width: 400, height: 200);
        Assert.True(addResult.Success, addResult.ErrorMessage);

        var result = _commands.SetNodeText(batch, slideIndex, shapeIndex: addResult.ShapeIndex!.Value, nodeIndex: 999, text: "x");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void AddNode_AppendsTopLevelNode_AndIncreasesCount()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 0, top: 0, width: 400, height: 200);
        Assert.True(addResult.Success, addResult.ErrorMessage);
        int shapeIndex = addResult.ShapeIndex!.Value;
        int before = addResult.NodeCount!.Value;

        var nodeResult = _commands.AddNode(batch, slideIndex, shapeIndex, text: "Extra Step");

        Assert.True(nodeResult.Success, nodeResult.ErrorMessage);
        Assert.Equal("Extra Step", nodeResult.NodeText);
        Assert.Equal(before + 1, nodeResult.NodeCount);

        var readBack = _commands.GetNodeText(batch, slideIndex, shapeIndex, nodeIndex: nodeResult.NodeIndex!.Value);
        Assert.True(readBack.Success, readBack.ErrorMessage);
        Assert.Equal("Extra Step", readBack.NodeText);
    }

    [Fact]
    public void AddChildNode_NestsUnderParent_AndIncreasesCount()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        // Organization Chart starts with a single root node — ideal for verifying hierarchy.
        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Organization Chart", left: 0, top: 0, width: 400, height: 300);
        Assert.True(addResult.Success, addResult.ErrorMessage);
        int shapeIndex = addResult.ShapeIndex!.Value;
        int before = addResult.NodeCount!.Value;

        var childResult = _commands.AddChildNode(batch, slideIndex, shapeIndex, parentNodeIndex: 1, text: "VP Engineering");

        Assert.True(childResult.Success, childResult.ErrorMessage);
        Assert.Equal("VP Engineering", childResult.NodeText);
        Assert.Equal(before + 1, childResult.NodeCount);

        var readBack = _commands.GetNodeText(batch, slideIndex, shapeIndex, nodeIndex: childResult.NodeIndex!.Value);
        Assert.True(readBack.Success, readBack.ErrorMessage);
        Assert.Equal("VP Engineering", readBack.NodeText);
    }

    [Fact]
    public void DeleteNode_RemovesNode_AndDecreasesCount()
    {
        var batch = _fixture.Batch;
        int slideIndex = AddIsolatedSlide();

        var addResult = _commands.AddSmartArt(batch, slideIndex, layoutName: "Basic Process", left: 0, top: 0, width: 400, height: 200);
        Assert.True(addResult.Success, addResult.ErrorMessage);
        int shapeIndex = addResult.ShapeIndex!.Value;

        var nodeResult = _commands.AddNode(batch, slideIndex, shapeIndex, text: "Temp Node");
        Assert.True(nodeResult.Success, nodeResult.ErrorMessage);
        int countAfterAdd = nodeResult.NodeCount!.Value;

        var deleteResult = _commands.DeleteNode(batch, slideIndex, shapeIndex, nodeIndex: nodeResult.NodeIndex!.Value);

        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);
        Assert.Equal(countAfterAdd - 1, deleteResult.NodeCount);
    }
}
