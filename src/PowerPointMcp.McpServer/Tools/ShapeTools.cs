using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Shape tools: add rectangles/text boxes, count, delete, and reposition/resize shapes on a slide.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="ShapeCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class ShapeTools
{
    private static readonly ShapeCommands Commands = new();

    /// <summary>Adds a rectangle shape to the given slide.</summary>
    [McpServerTool(Name = "add_rectangle")]
    [Description("Add a rectangle shape to the given slide.")]
    public static string AddRectangle(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the rectangle to.")] int slideIndex,
        [Description("Left position in points.")] float left,
        [Description("Top position in points.")] float top,
        [Description("Width in points.")] float width,
        [Description("Height in points.")] float height,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_rectangle", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddRectangle(batch, slideIndex, left, top, width, height));
        });

    /// <summary>Adds a text box with the given text to the given slide.</summary>
    [McpServerTool(Name = "add_text_box")]
    [Description("Add a text box with the given text to the given slide.")]
    public static string AddTextBox(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the text box to.")] int slideIndex,
        [Description("Left position in points.")] float left,
        [Description("Top position in points.")] float top,
        [Description("Width in points.")] float width,
        [Description("Height in points.")] float height,
        [Description("Text content of the text box.")] string text,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_text_box", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.AddTextBox(batch, slideIndex, left, top, width, height, text));
        });

    /// <summary>Gets the number of shapes on the given slide.</summary>
    [McpServerTool(Name = "get_shape_count")]
    [Description("Get the number of shapes on the given slide.")]
    public static string GetShapeCount(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to inspect.")] int slideIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_shape_count", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetCount(batch, slideIndex));
        });

    /// <summary>Deletes the shape at the given 1-based index on the given slide.</summary>
    [McpServerTool(Name = "delete_shape")]
    [Description("Delete the shape at the given 1-based index on the given slide.")]
    public static string DeleteShape(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape to delete.")] int shapeIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("delete_shape", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.Delete(batch, slideIndex, shapeIndex));
        });

    /// <summary>Sets the position of a shape on the given slide.</summary>
    [McpServerTool(Name = "set_shape_position")]
    [Description("Set the position of a shape on the given slide.")]
    public static string SetShapePosition(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape to reposition.")] int shapeIndex,
        [Description("New left position in points.")] float left,
        [Description("New top position in points.")] float top,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_shape_position", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetPosition(batch, slideIndex, shapeIndex, left, top));
        });

    /// <summary>Sets the size of a shape on the given slide.</summary>
    [McpServerTool(Name = "set_shape_size")]
    [Description("Set the size of a shape on the given slide.")]
    public static string SetShapeSize(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide containing the shape.")] int slideIndex,
        [Description("1-based index of the shape to resize.")] int shapeIndex,
        [Description("New width in points.")] float width,
        [Description("New height in points.")] float height,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_shape_size", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetSize(batch, slideIndex, shapeIndex, width, height));
        });

    private static string SerializeResult(ShapeOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            shapeIndex = result.ShapeIndex,
            shapeCount = result.ShapeCount,
            left = result.Left,
            top = result.Top,
            width = result.Width,
            height = result.Height,
            isError = result.Success ? (bool?)null : true
        });
}
