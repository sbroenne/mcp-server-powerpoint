using System.ComponentModel;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Single action-dispatch MCP tool for presentation lifecycle, template application, and
/// document-property operations: create a file, open/close a session, save, list open
/// sessions, apply a template, and read/write document properties.
/// </summary>
/// <remarks>
/// Mirrors mcp-server-excel's <c>ExcelFileTool</c> shape (one hand-written tool named
/// "presentation" with an <see cref="PresentationToolAction"/> enum parameter and an OPTIONAL
/// <c>sessionId</c>, since <c>create</c>/<c>open</c> establish a session rather than requiring
/// one) instead of exposing one MCP tool per verb. Collapsed from 12 individually-named tools
/// (create_presentation, open_presentation, ...) — see .squad/decisions.md for the rationale:
/// PowerPoint's per-domain generator always requires a non-nullable sessionId, which doesn't fit
/// session-establishing actions, so this domain stays hand-written like Excel's file tool.
///
/// The <see cref="PresentationSessionRegistry"/> singleton is resolved from DI and injected into
/// the tool method by the MCP SDK (parameters not part of the JSON schema are satisfied from the
/// host's service provider). Tools stay thin: they marshal to Core commands and serialize the
/// result — no domain logic lives here.
/// </remarks>
[McpServerToolType]
public static class PresentationTools
{
    private static readonly PresentationCommands Commands = new();

    /// <summary>
    /// Presentation lifecycle, template, and document-property operations for an already-open
    /// or about-to-be-opened presentation.
    /// </summary>
    [McpServerTool(Name = "presentation")]
    [Description("Presentation lifecycle (create, open, save, close, list sessions), template restyling (apply-template, get-theme-name), and document-property (built-in and custom) operations. Actions: create, open, save, close, list, apply-template, get-theme-name, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property.")]
    public static string Presentation(
        [Description("The action to perform. One of: create, open, save, close, list, apply-template, get-theme-name, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property.")] PresentationToolAction action,
        [Description("Full Windows path to the presentation file. Required for: create (new .pptx/.pptm file; containing directory must already exist), open (existing .pptx/.pptm/.ppt file).")] string? filePath = null,
        [Description("The sessionId returned by create or open. Required for: save, close, apply-template, get-theme-name, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property.")] string? sessionId = null,
        [Description("Set true only when creating a macro-enabled .pptm file. Default: false. Used for: create.")] bool isMacroEnabled = false,
        [Description("Full Windows path to a .potx/.potm/.pot template file (a .pptx/.pptm presentation may also be used as a template source). Required for: apply-template.")] string? templatePath = null,
        [Description("Document property name. For set/get-document-property, one of: Title, Subject, Author, Keywords, Comments, Category, Manager, Company (case-insensitive). For custom-property actions, any user-defined name. Required for: set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property.")] string? propertyName = null,
        [Description("The new property value. Required for: set-document-property, set-custom-property.")] string? value = null,
        PresentationSessionRegistry? registry = null)
        => PowerPointToolsBase.ExecuteToolAction("presentation", action.ToActionString(), () =>
        {
            var reg = registry!;
            return action switch
            {
                PresentationToolAction.Create => HandleCreate(filePath, isMacroEnabled, reg),
                PresentationToolAction.Open => HandleOpen(filePath, reg),
                PresentationToolAction.Save => HandleSave(sessionId, reg),
                PresentationToolAction.Close => HandleClose(sessionId, reg),
                PresentationToolAction.List => HandleList(reg),
                PresentationToolAction.ApplyTemplate => HandleApplyTemplate(sessionId, templatePath, reg),
                PresentationToolAction.GetThemeName => HandleGetThemeName(sessionId, reg),
                PresentationToolAction.SetDocumentProperty => HandleSetDocumentProperty(sessionId, propertyName, value, reg),
                PresentationToolAction.GetDocumentProperty => HandleGetDocumentProperty(sessionId, propertyName, reg),
                PresentationToolAction.SetCustomProperty => HandleSetCustomProperty(sessionId, propertyName, value, reg),
                PresentationToolAction.GetCustomProperty => HandleGetCustomProperty(sessionId, propertyName, reg),
                PresentationToolAction.RemoveCustomProperty => HandleRemoveCustomProperty(sessionId, propertyName, reg),
                _ => PowerPointToolsBase.ValidationError($"Unknown action: {action}")
            };
        });

    /// <summary>
    /// Creates a new, empty PowerPoint presentation, saves it to disk, and leaves the session
    /// OPEN — returns a sessionId immediately. No synchronous dispose happens here, so the call
    /// cannot block on PowerPoint's slow shutdown sequence (see
    /// .squad/decisions/inbox/ripley-create-presentation-blocks-on-dispose.md).
    /// </summary>
    private static string HandleCreate(string? filePath, bool isMacroEnabled, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return PowerPointToolsBase.ValidationError("filePath is required for action=create.");
        }

        var sessionId = registry.Create(filePath);

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
            message = "Presentation created and saved; session left open. Use the returned sessionId with other actions, then action=close when finished."
        });
    }

    /// <summary>
    /// Opens an existing presentation and returns a session id used by all subsequent actions.
    /// </summary>
    private static string HandleOpen(string? filePath, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return PowerPointToolsBase.ValidationError("filePath is required for action=open.");
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
    }

    /// <summary>
    /// Saves the presentation associated with the given session.
    /// </summary>
    private static string HandleSave(string? sessionId, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return PowerPointToolsBase.ValidationError("sessionId is required for action=save.");
        }

        if (!registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        var result = Commands.Save(batch);
        return SerializeResult(result);
    }

    /// <summary>
    /// Closes a session: removes it from the registry immediately and starts disposing its batch
    /// (releasing the underlying PowerPoint process) on a background task.
    /// </summary>
    /// <remarks>
    /// PowerPoint's own post-Quit cleanup can legitimately take up to ~150-210s (bounded grace
    /// period + force-kill safety net — see .squad/decisions.md, Parker's shutdown hardening).
    /// This does NOT wait for that; it returns as soon as the session is removed from the
    /// registry, so the MCP client is never blocked. The host still guarantees the PowerPoint
    /// process is fully cleaned up before it exits (see
    /// <see cref="PresentationSessionRegistry.DisposeAll()"/>).
    /// </remarks>
    private static string HandleClose(string? sessionId, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return PowerPointToolsBase.ValidationError("sessionId is required for action=close.");
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
    }

    /// <summary>
    /// Lists all currently open presentation sessions.
    /// </summary>
    private static string HandleList(PresentationSessionRegistry registry)
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
    }

    /// <summary>
    /// Applies a PowerPoint template's masters/theme/layouts to the open presentation, preserving
    /// slide content.
    /// </summary>
    private static string HandleApplyTemplate(string? sessionId, string? templatePath, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(templatePath))
        {
            return PowerPointToolsBase.ValidationError("templatePath is required for action=apply-template.");
        }

        return SerializeResult(Commands.ApplyTemplate(batch, templatePath));
    }

    /// <summary>
    /// Reads the design/theme name currently applied to the open presentation.
    /// </summary>
    private static string HandleGetThemeName(string? sessionId, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        return SerializeResult(Commands.GetThemeName(batch));
    }

    /// <summary>
    /// Sets a built-in document metadata property (Title, Subject, Author, Keywords, Comments,
    /// Category, Manager, or Company) on the open presentation.
    /// </summary>
    private static string HandleSetDocumentProperty(string? sessionId, string? propertyName, string? value, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=set-document-property.");
        }

        return SerializeResult(Commands.SetDocumentProperty(batch, propertyName, value ?? string.Empty));
    }

    /// <summary>
    /// Reads a built-in document metadata property from the open presentation.
    /// </summary>
    private static string HandleGetDocumentProperty(string? sessionId, string? propertyName, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=get-document-property.");
        }

        return SerializeResult(Commands.GetDocumentProperty(batch, propertyName));
    }

    /// <summary>
    /// Creates or updates a custom (user-defined) string document property on the open
    /// presentation.
    /// </summary>
    private static string HandleSetCustomProperty(string? sessionId, string? propertyName, string? value, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=set-custom-property.");
        }

        return SerializeResult(Commands.SetCustomProperty(batch, propertyName, value ?? string.Empty));
    }

    /// <summary>
    /// Reads a custom (user-defined) document property from the open presentation.
    /// </summary>
    private static string HandleGetCustomProperty(string? sessionId, string? propertyName, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=get-custom-property.");
        }

        return SerializeResult(Commands.GetCustomProperty(batch, propertyName));
    }

    /// <summary>
    /// Removes a custom (user-defined) document property from the open presentation.
    /// </summary>
    private static string HandleRemoveCustomProperty(string? sessionId, string? propertyName, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=remove-custom-property.");
        }

        return SerializeResult(Commands.RemoveCustomProperty(batch, propertyName));
    }

    private static string SerializeResult(PresentationOperationResult result)
    {
        if (result.Success)
        {
            return PowerPointToolsBase.Serialize(new
            {
                success = true,
                presentationPath = result.PresentationPath,
                themeName = result.ThemeName,
                propertyName = result.PropertyName,
                propertyValue = result.PropertyValue
            });
        }

        return PowerPointToolsBase.Serialize(new
        {
            success = false,
            errorMessage = result.ErrorMessage,
            presentationPath = result.PresentationPath,
            themeName = result.ThemeName,
            propertyName = result.PropertyName,
            propertyValue = result.PropertyValue,
            isError = true
        });
    }
}
