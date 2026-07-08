using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Export;

/// <summary>
/// Export commands: render presentation slides to raster image files.
/// Operates within an already-open <see cref="IPresentationBatch"/>.
/// </summary>
/// <remarks>
/// Uses the PowerPoint COM <c>Slide.Export</c> and <c>Presentation.Export</c> APIs, which
/// are the native mechanisms for slide→image conversion and do not require any additional
/// rendering library. This is the visual-verification ("verify visually") differentiator
/// over competing PowerPoint MCP servers that cannot produce images from slides.
/// </remarks>
[ServiceCategory("export", "Export")]
[McpTool("export", Title = "Export Operations", Destructive = true, Category = "content",
    Description = "Export slides in an open presentation session to raster image files.")]
public interface IExportCommands
{
    /// <summary>
    /// Exports a single slide to an image file using PowerPoint's native COM rendering
    /// (<c>Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)</c>).
    /// </summary>
    /// <param name="batch">The open presentation batch to operate on.</param>
    /// <param name="slideIndex">1-based index of the slide to export.</param>
    /// <param name="outputPath">Full path for the output image file (e.g. <c>C:\output\slide1.png</c>).</param>
    /// <param name="format">
    /// PowerPoint filter name for the image format (e.g. <c>"PNG"</c>, <c>"JPG"</c>, <c>"GIF"</c>).
    /// Defaults to <c>"PNG"</c>.
    /// </param>
    /// <param name="width">Optional output width in pixels; 0 or null uses PowerPoint's default.</param>
    /// <param name="height">Optional output height in pixels; 0 or null uses PowerPoint's default.</param>
    /// <returns>
    /// Success result with <see cref="ExportOperationResult.ExportedFilePath"/> set; or a failure
    /// result with <see cref="ExportOperationResult.ErrorMessage"/> for bad input (out-of-range
    /// slide index, inaccessible output path). Unexpected COM exceptions propagate (Rule 1b).
    /// </returns>
    ExportOperationResult ExportSlideToImage(
        IPresentationBatch batch,
        int slideIndex,
        string outputPath,
        string format = "PNG",
        int? width = null,
        int? height = null);

    /// <summary>
    /// Exports every slide in the presentation to image files in the specified directory,
    /// using PowerPoint's native <c>Presentation.Export(Path, FilterName, ScaleWidth, ScaleHeight)</c>
    /// which renders all slides in a single COM call.
    /// </summary>
    /// <param name="batch">The open presentation batch to operate on.</param>
    /// <param name="outputDirectory">
    /// Directory where slide images will be written. Created if it does not exist.
    /// PowerPoint names the output files <c>Slide1.{ext}</c>, <c>Slide2.{ext}</c>, etc.
    /// </param>
    /// <param name="format">
    /// PowerPoint filter name for the image format (e.g. <c>"PNG"</c>, <c>"JPG"</c>).
    /// Defaults to <c>"PNG"</c>.
    /// </param>
    /// <returns>
    /// Success result with <see cref="ExportOperationResult.ExportedFilePaths"/> listing the
    /// generated files in slide order, and <see cref="ExportOperationResult.SlideCount"/> set;
    /// or a failure result for bad input. Unexpected COM exceptions propagate (Rule 1b).
    /// </returns>
    ExportOperationResult ExportAllSlidesToImages(
        IPresentationBatch batch,
        string outputDirectory,
        string format = "PNG");
}
