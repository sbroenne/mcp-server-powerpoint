using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Export;

/// <inheritdoc cref="IExportCommands"/>
public sealed class ExportCommands : IExportCommands
{
    /// <inheritdoc/>
    public ExportOperationResult ExportSlideToImage(
        IPresentationBatch batch,
        int slideIndex,
        string outputPath,
        string format = "PNG",
        int? width = null,
        int? height = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(outputPath);
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return new ExportOperationResult
            {
                Success = false,
                ErrorMessage = "Output path must not be empty."
            };
        }

        // Resolve to a full path up-front so error messages are unambiguous.
        string fullOutputPath = Path.GetFullPath(outputPath);
        string? outputDir = Path.GetDirectoryName(fullOutputPath);
        if (string.IsNullOrEmpty(outputDir))
        {
            return new ExportOperationResult
            {
                Success = false,
                ErrorMessage = $"Cannot determine output directory from path: {fullOutputPath}."
            };
        }

        // Create the output directory if it does not exist — graceful input handling (Rule 1b).
        // A DirectoryNotFoundException from PowerPoint COM would be an unexpected failure;
        // creating it here makes the bad-input case explicit and testable.
        try
        {
            Directory.CreateDirectory(outputDir);
        }
        catch (Exception ex)
        {
            return new ExportOperationResult
            {
                Success = false,
                ErrorMessage = $"Could not create output directory '{outputDir}': {ex.Message}"
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new ExportOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. " +
                                   $"The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            // Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)
            // ScaleWidth/ScaleHeight of 0 instructs PowerPoint to use its default slide dimensions.
            // No MsoTriState parameters here — using the typed PowerPoint.Slide interface directly.
            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            slide.Export(fullOutputPath, format, width ?? 0, height ?? 0);

            return new ExportOperationResult
            {
                Success = true,
                ExportedFilePath = fullOutputPath,
                SlideCount = 1
            };
        });
    }

    /// <inheritdoc/>
    public ExportOperationResult ExportAllSlidesToImages(
        IPresentationBatch batch,
        string outputDirectory,
        string format = "PNG")
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(outputDirectory);
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            return new ExportOperationResult
            {
                Success = false,
                ErrorMessage = "Output directory must not be empty."
            };
        }

        string fullOutputDir = Path.GetFullPath(outputDirectory);

        // Create the output directory before calling Presentation.Export — some PowerPoint
        // versions require the target directory to exist.
        try
        {
            Directory.CreateDirectory(fullOutputDir);
        }
        catch (Exception ex)
        {
            return new ExportOperationResult
            {
                Success = false,
                ErrorMessage = $"Could not create output directory '{fullOutputDir}': {ex.Message}"
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;

            // Presentation.Export(Path, FilterName, ScaleWidth, ScaleHeight)
            // Writes one image per slide to the directory, named Slide1.{ext}, Slide2.{ext}, …
            // A single COM call is more efficient than looping over Slide.Export per slide.
            // ScaleWidth/ScaleHeight = 0 → PowerPoint default dimensions.
            ctx.Presentation.Export(fullOutputDir, format, 0, 0);

            // Enumerate the generated files sorted so the list is in slide order.
            // We use the format as a case-insensitive extension filter; PowerPoint may emit
            // the extension in any case (PNG/png) depending on the Office version.
            string[] files = Directory.GetFiles(fullOutputDir, $"*.{format}", SearchOption.TopDirectoryOnly);
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);

            return new ExportOperationResult
            {
                Success = true,
                ExportedFilePaths = files,
                SlideCount = slideCount
            };
        });
    }
}
