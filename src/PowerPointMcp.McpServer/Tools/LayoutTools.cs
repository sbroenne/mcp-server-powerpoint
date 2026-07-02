using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Layout;
using Sbroenne.PowerPointMcp.McpServer.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Slide layout tools: apply/read a slide's built-in layout.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="LayoutCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class LayoutTools
{
    private static readonly LayoutCommands Commands = new();

    /// <summary>Applies a built-in slide layout by its PpSlideLayout enum member name.</summary>
    [McpServerTool(Name = "set_layout")]
    [Description("Apply a built-in slide layout by its PpSlideLayout enum member name (e.g. \"ppLayoutBlank\", \"ppLayoutTitleOnly\", \"ppLayoutText\").")]
    public static string SetLayout(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to apply the layout to.")] int slideIndex,
        [Description("PpSlideLayout enum member name, e.g. \"ppLayoutBlank\", \"ppLayoutTitleOnly\", \"ppLayoutText\".")] string layoutName,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_layout", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetLayout(batch, slideIndex, layoutName));
        });

    /// <summary>Gets the current slide's layout name.</summary>
    [McpServerTool(Name = "get_layout")]
    [Description("Get the current slide's layout name.")]
    public static string GetLayout(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to inspect.")] int slideIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_layout", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetLayout(batch, slideIndex));
        });

    private static string SerializeResult(LayoutOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            layoutName = result.LayoutName,
            isError = result.Success ? (bool?)null : true
        });
}
