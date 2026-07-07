using Sbroenne.PowerPointMcp.Core.Export;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for export commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Export")]
public class ExportCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly SlideCommands _slideCommands = new();
    private readonly ExportCommands _commands = new();

    public ExportCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ExportSlideToImage_ExportsSingleSlide_FileExistsAndNonEmpty()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string outputDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"export-{Guid.NewGuid():N}");
        string outputFile = Path.Combine(outputDir, "slide1.png");
        try
        {
            var result = _commands.ExportSlideToImage(batch, 1, outputFile);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(outputFile, result.ExportedFilePath, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(1, result.SlideCount);
            Assert.True(File.Exists(outputFile), "Exported image file must exist on disk.");
            Assert.True(new FileInfo(outputFile).Length > 0, "Exported image file must be non-empty.");
        }
        finally
        {
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void ExportSlideToImage_WithCustomDimensions_ProducesFile()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string outputDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"export-dim-{Guid.NewGuid():N}");
        string outputFile = Path.Combine(outputDir, "slide_800x600.png");
        try
        {
            var result = _commands.ExportSlideToImage(batch, 1, outputFile, "PNG", width: 800, height: 600);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.True(File.Exists(outputFile), "Exported image must exist for custom dimensions.");
            Assert.True(new FileInfo(outputFile).Length > 0);
        }
        finally
        {
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void ExportSlideToImage_WithInvalidIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string outputDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"export-err-{Guid.NewGuid():N}");
        try
        {
            var result = _commands.ExportSlideToImage(batch, 99, Path.Combine(outputDir, "slide99.png"));

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void ExportAllSlidesToImages_ExportsAllSlides_FilesExistAndNonEmpty()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string outputDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"export-all-{Guid.NewGuid():N}");
        try
        {
            // Add a second slide so we have 2 to export.
            _slideCommands.AddBlank(batch);
            _presentationCommands.Save(batch);

            var result = _commands.ExportAllSlidesToImages(batch, outputDir);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(2, result.SlideCount);
            Assert.NotNull(result.ExportedFilePaths);
            Assert.Equal(2, result.ExportedFilePaths!.Count);
            foreach (string file in result.ExportedFilePaths)
            {
                Assert.True(File.Exists(file), $"Expected exported file to exist: {file}");
                Assert.True(new FileInfo(file).Length > 0, $"Expected exported file to be non-empty: {file}");
            }
        }
        finally
        {
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void ExportAllSlidesToImages_CreatesOutputDirectory_WhenMissing()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        // Use a nested directory that does NOT exist yet.
        string outputDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"export-mkdir-{Guid.NewGuid():N}", "nested", "dir");
        try
        {
            Assert.False(Directory.Exists(outputDir), "Pre-condition: output directory must not exist.");

            var result = _commands.ExportAllSlidesToImages(batch, outputDir);

            Assert.True(result.Success);
            Assert.True(Directory.Exists(outputDir), "Output directory must be created by ExportAllSlidesToImages.");
            Assert.Equal(1, result.SlideCount);
        }
        finally
        {
            // Walk up to the top temp directory we created and remove it.
            string topDir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
            string exportRoot = outputDir;
            // Find the direct child of PowerPointMcpTests that we created.
            string[] parts = Path.GetRelativePath(topDir, exportRoot).Split(Path.DirectorySeparatorChar);
            if (parts.Length > 0)
            {
                string dirToDelete = Path.Combine(topDir, parts[0]);
                if (Directory.Exists(dirToDelete))
                    Directory.Delete(dirToDelete, recursive: true);
            }
        }
    }
}
