using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for shape commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Shape")]
public class ShapeCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _commands = new();

    public ShapeCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddRectangle_IncreasesShapeCount_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch, 1);
        Assert.Equal(1, countResult.ShapeCount);
    }

    [Fact]
    public void AddTextBox_SetsText_ReadableAfterReopen()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddTextBox(batch, 1, 0f, 0f, 200f, 40f, "Hello PowerPoint");
        Assert.True(result.Success);
        Assert.Equal(1, result.ShapeIndex);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        string text = batch.Execute((ctx, ct) =>
            ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Text);
        Assert.Equal("Hello PowerPoint", text);
    }

    [Fact]
    public void SetPositionAndSize_UpdatesShapeGeometry()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

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

    [Fact]
    public void Delete_RemovesShape_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        var deleteResult = _commands.Delete(batch, 1, 1);

        Assert.True(deleteResult.Success);
        Assert.Equal(0, deleteResult.ShapeCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch, 1);
        Assert.Equal(0, countResult.ShapeCount);
    }

    [Fact]
    public void AddRectangle_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddRectangle(batch, 99, 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddAutoShape_WithOval_IncreasesShapeCount_AndEchoesShapeTypeName()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeOval", 10f, 20f, 100f, 50f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal("msoShapeOval", result.ShapeTypeName);
    }

    [Fact]
    public void AddAutoShape_WithRightArrow_IncreasesShapeCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeRightArrow", 0f, 0f, 80f, 40f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeCount);
    }

    [Fact]
    public void AddAutoShape_WithUnrecognizedShapeType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeDoesNotExist", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddAutoShape_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 99, "msoShapeOval", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddLine_IncreasesShapeCount_AndEchoesCoordinates()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddLine(batch, 1, 10f, 20f, 200f, 100f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal(10f, result.BeginX);
        Assert.Equal(20f, result.BeginY);
        Assert.Equal(200f, result.EndX);
        Assert.Equal(100f, result.EndY);
    }

    [Fact]
    public void AddLine_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddLine(batch, 99, 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddConnector_WithStraightType_IncreasesShapeCount_AndEchoesTypeAndCoordinates()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorStraight", 5f, 5f, 150f, 75f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal("msoConnectorStraight", result.ConnectorTypeName);
        Assert.Equal(5f, result.BeginX);
        Assert.Equal(5f, result.BeginY);
        Assert.Equal(150f, result.EndX);
        Assert.Equal(75f, result.EndY);
    }

    [Fact]
    public void AddConnector_WithElbowType_IncreasesShapeCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorElbow", 0f, 0f, 100f, 100f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeCount);
    }

    [Fact]
    public void AddConnector_WithUnrecognizedConnectorType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorDoesNotExist", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddConnector_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 99, "msoConnectorStraight", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
