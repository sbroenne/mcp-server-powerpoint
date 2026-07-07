using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Speaker notes tools: set/get the notes text for a slide.
/// </summary>
/// <remarks>
/// Thin pass-through to <see cref="NotesCommands"/> — see <see cref="PresentationTools"/> for the
/// session → registry → Core command pattern this follows.
/// </remarks>
[McpServerToolType]
public static class NotesTools
{
    private static readonly NotesCommands Commands = new();

    /// <summary>Sets the speaker notes text for a slide.</summary>
    [McpServerTool(Name = "set_notes_text")]
    [Description("Set the speaker notes text for a slide.")]
    public static string SetNotesText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to set notes on.")] int slideIndex,
        [Description("The new speaker notes text.")] string text,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("set_notes_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.SetNotesText(batch, slideIndex, text));
        });

    /// <summary>Gets the speaker notes text for a slide.</summary>
    [McpServerTool(Name = "get_notes_text")]
    [Description("Get the speaker notes text for a slide.")]
    public static string GetNotesText(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to read notes from.")] int slideIndex,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("get_notes_text", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return SerializeResult(Commands.GetNotesText(batch, slideIndex));
        });

    private static string SerializeResult(NotesOperationResult result)
        => PowerPointToolsBase.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            notesText = result.NotesText,
            isError = result.Success ? (bool?)null : true
        });
}
