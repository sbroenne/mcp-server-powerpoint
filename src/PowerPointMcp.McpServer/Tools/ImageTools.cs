using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Image tools: embed a picture file into a slide.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="ImageCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class ImageTools
{
    private static readonly ImageCommands Commands = new();

    /// <summary>Adds a picture from a local file to the given slide, embedded into the presentation.</summary>
    [McpServerTool(Name = "add_picture")]
    [Description("Add a picture from a local file to the given slide, embedded (not linked) into the presentation.")]
    public static string AddPicture(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the picture to.")] int slideIndex,
        [Description("Full Windows path to the local image file (e.g. C:\\images\\logo.png).")] string imagePath,
        [Description("Left position in points.")] float left,
        [Description("Top position in points.")] float top,
        [Description("Width in points.")] float width,
        [Description("Height in points.")] float height,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_picture", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddPicture(batch, slideIndex, imagePath, left, top, width, height));
        });

    private static string SerializeResult(ImageOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            shapeIndex = result.ShapeIndex,
            shapeCount = result.ShapeCount,
            isError = result.Success ? (bool?)null : true
        });
}
