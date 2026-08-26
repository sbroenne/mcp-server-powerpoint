using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real Save As and Save Copy As tests using the presentation fixture.</summary>
public partial class PresentationPropertiesTests
{
    [Fact]
    public void SaveCopyAs_CreatesCopyAndLeavesSessionPathUnchanged()
    {
        string originalPath = _fixture.CreateFreshPresentation();
        string copyPath = CoreTestHelper.CreateUniqueTestFilePath();
        _fixture.TrackCreatedPath(copyPath);
        _fixture.Batch.Save();
        _fixture.Batch.Execute((ctx, ct) =>
        {
            ctx.Presentation.Slides.Add(
                2,
                Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
        });

        var result = _commands.SaveCopyAs(_fixture.Batch, copyPath);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(File.Exists(originalPath));
        Assert.True(File.Exists(copyPath));
        Assert.Equal(Path.GetFullPath(copyPath), result.PresentationPath, ignoreCase: true);
        Assert.Equal(Path.GetFullPath(originalPath), _fixture.Batch.PresentationPath, ignoreCase: true);

        _fixture.ReopenPresentation(copyPath);
        int slideCount = _fixture.Batch.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
        Assert.Equal(2, slideCount);
    }

    [Fact]
    public void SaveAs_FormatExtensionMismatch_ReturnsFailureAndPreservesSessionPath()
    {
        string originalPath = _fixture.CreateFreshPresentation();
        string mismatchedPath = Path.ChangeExtension(
            CoreTestHelper.CreateUniqueTestFilePath(),
            ".pptm");
        _fixture.TrackCreatedPath(mismatchedPath);

        var result = _commands.SaveAs(
            _fixture.Batch,
            mismatchedPath,
            PresentationSaveFormat.Pptx);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.False(File.Exists(mismatchedPath));
        Assert.Equal(Path.GetFullPath(originalPath), _fixture.Batch.PresentationPath, ignoreCase: true);
    }

    [Fact]
    public void SaveCopyAs_FormatExtensionMismatch_ReturnsFailureAndPreservesSessionPath()
    {
        string originalPath = _fixture.CreateFreshPresentation();
        string mismatchedPath = Path.ChangeExtension(
            CoreTestHelper.CreateUniqueTestFilePath(),
            ".pptm");
        _fixture.TrackCreatedPath(mismatchedPath);

        var result = _commands.SaveCopyAs(_fixture.Batch, mismatchedPath);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.False(File.Exists(mismatchedPath));
        Assert.Equal(Path.GetFullPath(originalPath), _fixture.Batch.PresentationPath, ignoreCase: true);
    }

    [Fact]
    public void SaveAs_ExistingDestinationWithoutOverwrite_ReturnsFailureAndPreservesSessionPath()
    {
        string originalPath = _fixture.CreateFreshPresentation();
        string existingPath = CoreTestHelper.CreateUniqueTestFilePath();
        File.WriteAllText(existingPath, "existing destination");
        _fixture.TrackCreatedPath(existingPath);

        var result = _commands.SaveAs(_fixture.Batch, existingPath);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("existing destination", File.ReadAllText(existingPath));
        Assert.Equal(Path.GetFullPath(originalPath), _fixture.Batch.PresentationPath, ignoreCase: true);
    }

    [Fact]
    public void SaveCopyAs_ExistingDestinationWithoutOverwrite_ReturnsFailure()
    {
        string originalPath = _fixture.CreateFreshPresentation();
        string existingPath = CoreTestHelper.CreateUniqueTestFilePath();
        File.WriteAllText(existingPath, "existing destination");
        _fixture.TrackCreatedPath(existingPath);

        var result = _commands.SaveCopyAs(_fixture.Batch, existingPath);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("existing destination", File.ReadAllText(existingPath));
        Assert.Equal(Path.GetFullPath(originalPath), _fixture.Batch.PresentationPath, ignoreCase: true);
    }

    [Fact]
    public void SaveAsAndSaveCopyAs_WithOverwrite_ReplaceExistingDestinations()
    {
        _fixture.CreateFreshPresentation();
        string saveAsPath = CoreTestHelper.CreateUniqueTestFilePath();
        string copyPath = CoreTestHelper.CreateUniqueTestFilePath();
        File.WriteAllText(saveAsPath, "existing save-as destination");
        File.WriteAllText(copyPath, "existing copy destination");
        _fixture.TrackCreatedPath(saveAsPath);
        _fixture.TrackCreatedPath(copyPath);

        var copyResult = _commands.SaveCopyAs(
            _fixture.Batch,
            copyPath,
            overwrite: true);
        Assert.True(copyResult.Success);
        _fixture.ReopenPresentation(copyPath);
        Assert.Equal(
            1,
            _fixture.Batch.Execute((ctx, ct) => ctx.Presentation.Slides.Count));

        _fixture.CreateFreshPresentation();
        var saveAsResult = _commands.SaveAs(
            _fixture.Batch,
            saveAsPath,
            PresentationSaveFormat.Pptx,
            overwrite: true);

        Assert.True(saveAsResult.Success);
        Assert.Equal(Path.GetFullPath(saveAsPath), _fixture.Batch.PresentationPath, ignoreCase: true);
        _fixture.ReopenPresentation(saveAsPath);
        Assert.Equal(
            1,
            _fixture.Batch.Execute((ctx, ct) => ctx.Presentation.Slides.Count));
    }
}
