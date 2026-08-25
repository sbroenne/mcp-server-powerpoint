using Sbroenne.PowerPointMcp.Core.Accessibility;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real PowerPoint integration tests for deterministic accessibility operations.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Accessibility")]
public sealed class AccessibilityCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly AccessibilityCommands _commands = new();
    private readonly ImageCommands _imageCommands = new();
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _shapeCommands = new();

    public AccessibilityCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Audit_FindsPictureWithoutAlternativeText()
    {
        _fixture.CreateFreshPresentation();
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _imageCommands.AddPicture(_fixture.Batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var result = _commands.Audit(_fixture.Batch);

            Assert.True(result.Success, result.ErrorMessage);
            var issue = Assert.Single(result.Issues!, i => i.Code == "missing-alt-text");
            Assert.Equal(1, issue.SlideIndex);
            Assert.Equal(1, issue.ShapeIndex);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void Audit_CleanBlankPresentation_ReturnsNoIssues()
    {
        _fixture.CreateFreshPresentation();

        var result = _commands.Audit(_fixture.Batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.Issues);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void SetReadingOrder_ThenGetReadingOrder_RoundTripsPermutation()
    {
        _fixture.CreateFreshPresentation();
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 10f, 100f, 30f, "First");
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 50f, 100f, 30f, "Second");
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 90f, 100f, 30f, "Third");

        var setResult = _commands.SetReadingOrder(_fixture.Batch, 1, [3, 1, 2]);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Null(setResult.ErrorMessage);
        Assert.Equal([3, 1, 2], setResult.ReadingOrder);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        var getResult = _commands.GetReadingOrder(_fixture.Batch, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal([3, 1, 2], getResult.ReadingOrder);
    }

    [Fact]
    public void SetReadingOrder_WithDuplicateAndMissingShapeIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 10f, 100f, 30f, "First");
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 50f, 100f, 30f, "Second");
        _shapeCommands.AddTextBox(_fixture.Batch, 1, 10f, 90f, 100f, 30f, "Third");

        var result = _commands.SetReadingOrder(_fixture.Batch, 1, [1, 1, 3]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
    }
}
