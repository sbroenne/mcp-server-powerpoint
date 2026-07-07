using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Slide lifecycle tools: add, count, and delete slides within an open session.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="SlideCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class SlideTools
{
    private static readonly SlideCommands Commands = new();

    /// <summary>Adds a new blank slide at the end of the presentation.</summary>
    [McpServerTool(Name = "add_slide")]
    [Description("Add a new blank slide at the end of the presentation.")]
    public static string AddSlide(
        [Description("The session id returned by open_presentation.")] string sessionId,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_slide", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddBlank(batch));
        });

    /// <summary>Gets the current number of slides in the presentation.</summary>
    [McpServerTool(Name = "get_slide_count")]
    [Description("Get the current number of slides in the presentation.")]
    public static string GetSlideCount(
        [Description("The session id returned by open_presentation.")] string sessionId,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_slide_count", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetCount(batch));
        });

    /// <summary>Deletes the slide at the given 1-based index.</summary>
    [McpServerTool(Name = "delete_slide")]
    [Description("Delete the slide at the given 1-based index.")]
    public static string DeleteSlide(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to delete.")] int slideIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("delete_slide", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.Delete(batch, slideIndex));
        });

    private static string SerializeResult(SlideOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            slideIndex = result.SlideIndex,
            slideCount = result.SlideCount,
            isError = result.Success ? (bool?)null : true
        });
}
