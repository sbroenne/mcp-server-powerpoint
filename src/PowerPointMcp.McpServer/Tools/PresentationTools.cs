using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.McpServer.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Presentation lifecycle tools: create a file, open/close a session, save, and list open
/// sessions. This is the first hand-written vertical slice of the MCP surface — it proves the
/// session → registry → Core command pipeline end-to-end.
/// </summary>
/// <remarks>
/// The <see cref="PresentationSessionRegistry"/> singleton is resolved from DI and injected into
/// each tool method by the MCP SDK (parameters not part of the JSON schema are satisfied from the
/// host's service provider). Tools stay thin: they marshal to Core commands and serialize the
/// result — no domain logic lives here.
/// </remarks>
[McpServerToolType]
public static class PresentationTools
{
    private static readonly PresentationCommands Commands = new();

    /// <summary>
    /// Creates a new, empty PowerPoint presentation, saves it to disk, and leaves the session
    /// OPEN — returns a sessionId immediately, exactly like <c>open_presentation</c>. No
    /// synchronous dispose happens here, so the call cannot block on PowerPoint's slow shutdown
    /// sequence (see .squad/decisions/inbox/ripley-create-presentation-blocks-on-dispose.md).
    /// </summary>
    [McpServerTool(Name = "create_presentation")]
    [Description("Create a new empty PowerPoint presentation file on disk and leave it OPEN. The containing directory must already exist. Returns a sessionId that must be passed to subsequent tools (save_presentation, close_presentation, slide/shape tools, etc.) — there is no separate open_presentation call needed for a freshly created file. Returns immediately; does not wait for any PowerPoint shutdown.")]
    public static string CreatePresentation(
        [Description("Full Windows path to the new presentation. Use .pptx (standard) or .pptm (macro-enabled). Example: C:\\Users\\me\\Documents\\deck.pptx")] string filePath,
        [Description("Set true only when creating a macro-enabled .pptm file. Default: false.")] bool isMacroEnabled = false,
        PresentationSessionRegistry? registry = null)
        => PowerPointToolsBase.ExecuteToolAction("create_presentation", () =>
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return PowerPointToolsBase.ValidationError("filePath is required for create_presentation.");
            }

            var sessionId = registry!.Create(filePath);

            // Persist the new file to disk immediately through the still-open batch — no
            // Dispose(), so this cannot block on PowerPoint's shutdown/grace-period sequence.
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Session {sessionId} was created but could not be resolved.");
            }

            var result = Commands.Save(batch);
            if (!result.Success)
            {
                return SerializeResult(result);
            }

            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                sessionId,
                presentationPath = result.PresentationPath,
                message = "Presentation created and saved; session left open. Use the returned sessionId with other tools, then close_presentation when finished."
            });
        });

    /// <summary>
    /// Opens an existing presentation and returns a session id used by all subsequent tools.
    /// </summary>
    [McpServerTool(Name = "open_presentation")]
    [Description("Open an existing PowerPoint presentation and start a session. Returns a sessionId that must be passed to all subsequent tools (save_presentation, close_presentation, etc.).")]
    public static string OpenPresentation(
        [Description("Full Windows path to an existing .pptx, .pptm, or .ppt file. Example: C:\\Users\\me\\Documents\\deck.pptx")] string filePath,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("open_presentation", () =>
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return PowerPointToolsBase.ValidationError("filePath is required for open_presentation.");
            }

            if (!File.Exists(filePath))
            {
                return PowerPointToolsBase.ValidationError($"File not found: {filePath}");
            }

            var sessionId = registry.Open(filePath);
            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                sessionId,
                presentationPath = filePath
            });
        });

    /// <summary>
    /// Saves the presentation associated with the given session.
    /// </summary>
    [McpServerTool(Name = "save_presentation")]
    [Description("Save the presentation for an open session to its current file. Requires a sessionId from open_presentation.")]
    public static string SavePresentation(
        [Description("The sessionId returned by open_presentation.")] string sessionId,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("save_presentation", () =>
        {
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                return PowerPointToolsBase.ValidationError("sessionId is required for save_presentation.");
            }

            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            var result = Commands.Save(batch);
            return SerializeResult(result);
        });

    /// <summary>
    /// Closes a session: removes it from the registry immediately and starts disposing its batch
    /// (releasing the underlying PowerPoint process) on a background task.
    /// </summary>
    /// <remarks>
    /// PowerPoint's own post-Quit cleanup can legitimately take up to ~150-210s (bounded grace
    /// period + force-kill safety net — see .squad/decisions.md, Parker's shutdown hardening).
    /// This tool call does NOT wait for that; it returns as soon as the session is removed from
    /// the registry, so the MCP client is never blocked. The host still guarantees the PowerPoint
    /// process is fully cleaned up before it exits (see
    /// <see cref="PresentationSessionRegistry.DisposeAll()"/>).
    /// </remarks>
    [McpServerTool(Name = "close_presentation")]
    [Description("Close an open session and release its PowerPoint process. Call this when finished with a presentation. Save first with save_presentation if you need to persist changes. Returns immediately; PowerPoint shuts down in the background (can take up to a few minutes) and the session is already gone from list_sessions.")]
    public static string ClosePresentation(
        [Description("The sessionId returned by open_presentation.")] string sessionId,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("close_presentation", () =>
        {
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                return PowerPointToolsBase.ValidationError("sessionId is required for close_presentation.");
            }

            var closed = registry.Close(sessionId);
            if (!closed)
            {
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
            }

            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                sessionId,
                closed = true,
                message = "Session closed; PowerPoint is shutting down in the background."
            });
        });

    /// <summary>
    /// Lists all currently open presentation sessions.
    /// </summary>
    [McpServerTool(Name = "list_sessions")]
    [Description("List all currently open PowerPoint presentation sessions, including their id, file path, and whether the PowerPoint process is still alive.")]
    public static string ListSessions(PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("list_sessions", () =>
        {
            var sessions = registry.List()
                .Select(s => new
                {
                    sessionId = s.SessionId,
                    presentationPath = s.PresentationPath,
                    isPowerPointProcessAlive = s.IsPowerPointProcessAlive
                })
                .ToArray();

            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                count = sessions.Length,
                sessions
            });
        });

    private static string SerializeResult(PresentationOperationResult result)
    {
        if (result.Success)
        {
            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                presentationPath = result.PresentationPath
            });
        }

        return PowerPointToolsBase.Serialize(new
        {
            success = false,
            errorMessage = result.ErrorMessage,
            presentationPath = result.PresentationPath,
            isError = true
        });
    }
}
