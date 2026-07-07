using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Text frame tools: set/get text and basic font formatting (size, bold, color) for a shape's
/// text range.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="TextFrameCommands"/> — see <see cref="PresentationTools"/> for
/// the session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class TextFrameTools
{
    private static readonly TextFrameCommands Commands = new();

    /// <summary>Sets the text content of a shape's text frame.</summary>
    [McpServerTool(Name = "set_text")]
    [Description("Set the text content of a shape's text frame.")]
    public static string SetText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape whose text frame to update.")] int shapeIndex,
        [Description("The new text content.")] string text,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetText(batch, slideIndex, shapeIndex, text));
        });

    /// <summary>Gets the text content of a shape's text frame.</summary>
    [McpServerTool(Name = "get_text")]
    [Description("Get the text content of a shape's text frame.")]
    public static string GetText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape whose text frame to read.")] int shapeIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetText(batch, slideIndex, shapeIndex));
        });

    /// <summary>Sets the font size (in points) of a shape's entire text range.</summary>
    [McpServerTool(Name = "set_font_size")]
    [Description("Set the font size (in points) of a shape's entire text range.")]
    public static string SetFontSize(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape whose text to format.")] int shapeIndex,
        [Description("New font size in points.")] float fontSize,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_font_size", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetFontSize(batch, slideIndex, shapeIndex, fontSize));
        });

    /// <summary>Sets whether a shape's entire text range is bold.</summary>
    [McpServerTool(Name = "set_bold")]
    [Description("Set whether a shape's entire text range is bold.")]
    public static string SetBold(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape whose text to format.")] int shapeIndex,
        [Description("True to make the text bold, false to unbold it.")] bool bold,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_bold", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetBold(batch, slideIndex, shapeIndex, bold));
        });

    /// <summary>Sets the font color (RGB) of a shape's entire text range.</summary>
    [McpServerTool(Name = "set_font_color")]
    [Description("Set the font color (RGB) of a shape's entire text range.")]
    public static string SetFontColor(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape whose text to format.")] int shapeIndex,
        [Description("Red component (0-255).")] byte red,
        [Description("Green component (0-255).")] byte green,
        [Description("Blue component (0-255).")] byte blue,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_font_color", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetFontColor(batch, slideIndex, shapeIndex, red, green, blue));
        });

    private static string SerializeResult(TextFrameOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            text = result.Text,
            fontSize = result.FontSize,
            bold = result.Bold,
            colorRgb = result.ColorRgb,
            isError = result.Success ? (bool?)null : true
        });
}
