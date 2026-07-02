using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Export;
using Sbroenne.PowerPointMcp.McpServer.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Export tools: render presentation slides to raster image files using PowerPoint's native
/// COM rendering.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="ExportCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class ExportTools
{
    private static readonly ExportCommands Commands = new();

    /// <summary>Exports a single slide to an image file using PowerPoint's native COM rendering.</summary>
    [McpServerTool(Name = "export_slide_to_image")]
    [Description("Export a single slide to an image file using PowerPoint's native COM rendering.")]
    public static string ExportSlideToImage(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to export.")] int slideIndex,
        [Description("Full path for the output image file (e.g. C:\\output\\slide1.png).")] string outputPath,
        [Description("PowerPoint filter name for the image format (e.g. \"PNG\", \"JPG\", \"GIF\"). Defaults to \"PNG\".")] string format = "PNG",
        [Description("Optional output width in pixels; omit or 0 to use PowerPoint's default.")] int? width = null,
        [Description("Optional output height in pixels; omit or 0 to use PowerPoint's default.")] int? height = null,
        PresentationSessionRegistry? registry = null)
        => PowerPointToolsBase.ExecuteToolAction("export_slide_to_image", () =>
        {
            if (registry is null || !registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.ExportSlideToImage(batch, slideIndex, outputPath, format, width, height));
        });

    /// <summary>
    /// Exports every slide in the presentation to image files in the specified directory.
    /// </summary>
    [McpServerTool(Name = "export_all_slides_to_images")]
    [Description("Export every slide in the presentation to image files in the specified directory. PowerPoint names the output files Slide1.{ext}, Slide2.{ext}, etc.")]
    public static string ExportAllSlidesToImages(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("Directory where slide images will be written. Created if it does not exist.")] string outputDirectory,
        [Description("PowerPoint filter name for the image format (e.g. \"PNG\", \"JPG\"). Defaults to \"PNG\".")] string format = "PNG",
        PresentationSessionRegistry? registry = null)
        => PowerPointToolsBase.ExecuteToolAction("export_all_slides_to_images", () =>
        {
            if (registry is null || !registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.ExportAllSlidesToImages(batch, outputDirectory, format));
        });

    private static string SerializeResult(ExportOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            exportedFilePath = result.ExportedFilePath,
            exportedFilePaths = result.ExportedFilePaths,
            slideCount = result.SlideCount,
            isError = result.Success ? (bool?)null : true
        });
}
