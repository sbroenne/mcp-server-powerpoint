using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Single action-dispatch MCP tool for presentation lifecycle, template application, and
/// document-property and string-tag operations: create a file, open/close/test a session, list
/// open sessions, Save As/copy, apply a template, read/write document properties, manage the
/// advisory Mark as Final flag, and manage string tags.
/// </summary>
/// <remarks>
/// Mirrors mcp-server-excel's <c>ExcelFileTool</c> shape (one hand-written tool named
/// "presentation" with an <see cref="PresentationToolAction"/> enum parameter and an OPTIONAL
/// <c>presentation_session_id</c>, since <c>create</c>/<c>open</c> establish a session rather than requiring
/// one) instead of exposing one MCP tool per verb. PowerPoint's per-domain generator always
/// requires a non-nullable presentation_session_id, which doesn't fit
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
    /// Presentation lifecycle, Save As/copy, template, document-property, advisory Mark as Final,
    /// and string-tag operations for an already-open or about-to-be-opened presentation.
    /// </summary>
    [McpServerTool(
        Name = "presentation",
        UseStructuredContent = true,
        OutputSchemaType = typeof(PresentationToolOutputSchema))]
    [Description("Presentation lifecycle, Save As/copy, template, document-property, advisory Mark as Final, and string-tag operations. Mark as Final is not authentication, encryption, or access control. Actions: create, open, close, list, test, save-as, save-copy-as, apply-template, get-theme-name, get-final, set-final, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property, set-tag, get-tag, list-tags, delete-tag.")]
    public static Task<CallToolResult> Presentation(
        [Description("The action to perform. One of: create, open, close, list, test, save-as, save-copy-as, apply-template, get-theme-name, get-final, set-final, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property, set-tag, get-tag, list-tags, delete-tag.")] PresentationToolAction action,
        [Description("Full Windows path to the presentation file. Required for: create (new .pptx/.pptm file; containing directory must already exist), open or test (existing .pptx/.pptm/.ppt file).")] string? filePath = null,
        [Description("The presentation_session_id returned by create or open. Required for: close, save-as, save-copy-as, apply-template, get-theme-name, get-final, set-final, set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property, set-tag, get-tag, list-tags, delete-tag.")] string? presentation_session_id = null,
        [Description("Save the presentation before closing. Used for: close. Default: false.")] bool? save = null,
        [Description("Set true only when creating a macro-enabled .pptm file. Default: false. Used for: create.")] bool? isMacroEnabled = null,
        [Description("Full Windows destination path. Required for: save-as, save-copy-as. The containing directory must already exist.")] string? targetPath = null,
        [Description("Output format for save-as: auto, pptx, pptm, or ppt. Default: auto infers from targetPath.")] PresentationSaveFormat? format = null,
        [Description("Whether an existing destination file may be replaced. Used for: save-as, save-copy-as. Default: false.")] bool? overwrite = null,
        [Description("Full Windows path to a .potx/.potm/.pot template file (a .pptx/.pptm presentation may also be used as a template source). Required for: apply-template.")] string? templatePath = null,
        [Description("Set true to save current changes and mark the presentation as final, or false to clear the flag. Required for: set-final. Mark as Final is an advisory editing flag only; it is not authentication, encryption, or access control.")] bool? isFinal = null,
        [Description("Document property name. For set/get-document-property, one of: Title, Subject, Author, Keywords, Comments, Category, Manager, Company (case-insensitive). For custom-property actions, any user-defined name. Required for: set-document-property, get-document-property, set-custom-property, get-custom-property, remove-custom-property.")] string? propertyName = null,
        [Description("The new property value. Required for: set-document-property, set-custom-property.")] string? value = null,
        [Description("Case-insensitive string tag name. Letter casing is normalized to invariant uppercase; whitespace is preserved. Required for: set-tag, get-tag, delete-tag.")] string? tagName = null,
        [Description("String tag value, preserved exactly without case normalization. Required for: set-tag.")] string? tagValue = null,
        PresentationSessionRegistry? registry = null,
        CancellationToken cancellationToken = default)
        => PowerPointToolsBase.ExecuteToolActionAsync("presentation", action.ToActionString(), () =>
        {
            ValidateActionParameters(action, filePath, presentation_session_id, save, isMacroEnabled, targetPath, format, overwrite, templatePath, isFinal, propertyName, value, tagName, tagValue);
            var reg = registry!;
            return action switch
            {
                PresentationToolAction.Create => HandleCreate(filePath, isMacroEnabled == true, reg, cancellationToken),
                PresentationToolAction.Open => HandleOpen(filePath, reg, cancellationToken),
                PresentationToolAction.Close => HandleClose(presentation_session_id, save == true, reg),
                PresentationToolAction.List => HandleList(reg),
                PresentationToolAction.Test => HandleTest(filePath),
                PresentationToolAction.SaveAs => HandleSaveAs(presentation_session_id, targetPath, format, overwrite == true, reg),
                PresentationToolAction.SaveCopyAs => HandleSaveCopyAs(presentation_session_id, targetPath, overwrite == true, reg),
                PresentationToolAction.ApplyTemplate => HandleApplyTemplate(presentation_session_id, templatePath, reg),
                PresentationToolAction.GetThemeName => HandleGetThemeName(presentation_session_id, reg),
                PresentationToolAction.GetFinal => HandleGetFinal(presentation_session_id, reg),
                PresentationToolAction.SetFinal => HandleSetFinal(presentation_session_id, isFinal, reg),
                PresentationToolAction.SetDocumentProperty => HandleSetDocumentProperty(presentation_session_id, propertyName, value, reg),
                PresentationToolAction.GetDocumentProperty => HandleGetDocumentProperty(presentation_session_id, propertyName, reg),
                PresentationToolAction.SetCustomProperty => HandleSetCustomProperty(presentation_session_id, propertyName, value, reg),
                PresentationToolAction.GetCustomProperty => HandleGetCustomProperty(presentation_session_id, propertyName, reg),
                PresentationToolAction.RemoveCustomProperty => HandleRemoveCustomProperty(presentation_session_id, propertyName, reg),
                PresentationToolAction.SetTag => HandleSetTag(presentation_session_id, tagName, tagValue, reg),
                PresentationToolAction.GetTag => HandleGetTag(presentation_session_id, tagName, reg),
                PresentationToolAction.ListTags => HandleListTags(presentation_session_id, reg),
                PresentationToolAction.DeleteTag => HandleDeleteTag(presentation_session_id, tagName, reg),
                _ => PowerPointToolsBase.ValidationError($"Unknown action: {action}")
            };
        }, cancellationToken, registry, presentation_session_id);

    internal static void ValidateActionParameterNames(
        string action,
        IEnumerable<string> suppliedParameters)
    {
        var parsedAction = Enum.GetValues<PresentationToolAction>()
            .SingleOrDefault(value => string.Equals(value.ToActionString(), action, StringComparison.OrdinalIgnoreCase));
        if (!Enum.IsDefined(parsedAction))
            throw new ArgumentException($"Unknown action: {action}");

        var allowed = GetAllowedParameterNames(parsedAction);
        var invalid = suppliedParameters
            .Where(parameter => parameter != "action" && !allowed.Contains(parameter))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(parameter => parameter, StringComparer.Ordinal)
            .ToArray();
        if (invalid.Length > 0)
            throw new ArgumentException(
                $"Parameter(s) not valid for action '{action}': {string.Join(", ", invalid)}.");
    }

    private static void ValidateActionParameters(
        PresentationToolAction action,
        string? filePath,
        string? sessionId,
        bool? save,
        bool? isMacroEnabled,
        string? targetPath,
        PresentationSaveFormat? format,
        bool? overwrite,
        string? templatePath,
        bool? isFinal,
        string? propertyName,
        string? value,
        string? tagName,
        string? tagValue)
    {
        var supplied = new List<string>();
        if (filePath != null) supplied.Add("filePath");
        if (sessionId != null) supplied.Add("presentation_session_id");
        if (save != null) supplied.Add("save");
        if (isMacroEnabled != null) supplied.Add("isMacroEnabled");
        if (targetPath != null) supplied.Add("targetPath");
        if (format != null) supplied.Add("format");
        if (overwrite != null) supplied.Add("overwrite");
        if (templatePath != null) supplied.Add("templatePath");
        if (isFinal != null) supplied.Add("isFinal");
        if (propertyName != null) supplied.Add("propertyName");
        if (value != null) supplied.Add("value");
        if (tagName != null) supplied.Add("tagName");
        if (tagValue != null) supplied.Add("tagValue");

        var allowed = GetAllowedParameterNames(action);
        var inapplicable = supplied.Where(parameter => !allowed.Contains(parameter)).ToArray();
        if (inapplicable.Length > 0)
        {
            throw new ArgumentException(
                $"Parameter(s) not valid for action '{action.ToActionString()}': {string.Join(", ", inapplicable)}.");
        }
    }

    private static HashSet<string> GetAllowedParameterNames(PresentationToolAction action)
    {
        string[] allowedParameters = action switch
        {
            PresentationToolAction.Create => ["filePath", "isMacroEnabled"],
            PresentationToolAction.Open or PresentationToolAction.Test => ["filePath"],
            PresentationToolAction.Close => ["presentation_session_id", "save"],
            PresentationToolAction.SaveAs => ["presentation_session_id", "targetPath", "format", "overwrite"],
            PresentationToolAction.SaveCopyAs => ["presentation_session_id", "targetPath", "overwrite"],
            PresentationToolAction.ApplyTemplate => ["presentation_session_id", "templatePath"],
            PresentationToolAction.GetThemeName or PresentationToolAction.GetFinal => ["presentation_session_id"],
            PresentationToolAction.SetFinal => ["presentation_session_id", "isFinal"],
            PresentationToolAction.SetDocumentProperty or PresentationToolAction.SetCustomProperty => ["presentation_session_id", "propertyName", "value"],
            PresentationToolAction.GetDocumentProperty or PresentationToolAction.GetCustomProperty or PresentationToolAction.RemoveCustomProperty => ["presentation_session_id", "propertyName"],
            PresentationToolAction.SetTag => ["presentation_session_id", "tagName", "tagValue"],
            PresentationToolAction.GetTag or PresentationToolAction.DeleteTag => ["presentation_session_id", "tagName"],
            PresentationToolAction.ListTags => ["presentation_session_id"],
            _ => []
        };
        return new HashSet<string>(allowedParameters, StringComparer.Ordinal);
    }

    /// <summary>
    /// Creates a new, empty PowerPoint presentation, saves it to disk, and leaves the session
    /// OPEN — returns a presentation_session_id immediately. No synchronous dispose happens here, so the call
    /// cannot block on PowerPoint's slow shutdown sequence.
    /// </summary>
    private static string HandleCreate(
        string? filePath, bool isMacroEnabled, PresentationSessionRegistry registry, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return PowerPointToolsBase.ValidationError("filePath is required for action=create.");
        }

        if (isMacroEnabled && !Path.GetExtension(filePath).Equals(".pptm", StringComparison.OrdinalIgnoreCase))
        {
            return PowerPointToolsBase.ValidationError("isMacroEnabled=true requires a .pptm file path.");
        }

        var sessionId = registry.Create(filePath);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
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
                presentation_session_id = sessionId,
                presentationPath = result.PresentationPath,
                message = "Presentation created and saved; session left open. Use the returned presentation_session_id with other actions, then action=close when finished."
            });
        }
        finally
        {
            if (cancellationToken.IsCancellationRequested)
                registry.Close(sessionId);
        }
    }

    /// <summary>
    /// Opens an existing presentation and returns a session id used by all subsequent actions.
    /// </summary>
    private static string HandleOpen(
        string? filePath, PresentationSessionRegistry registry, CancellationToken cancellationToken)
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
        if (cancellationToken.IsCancellationRequested)
        {
            registry.Close(sessionId);
            cancellationToken.ThrowIfCancellationRequested();
        }
        return PowerPointToolsBase.Serialize(new
        {
            success = true,
            presentation_session_id = sessionId,
            presentationPath = filePath
        });
    }

    /// <summary>
    /// Closes a session: removes it from the registry immediately and starts disposing its batch
    /// (releasing the underlying PowerPoint process) on a background task.
    /// </summary>
    /// <remarks>
    /// PowerPoint's own post-Quit cleanup can legitimately take up to ~150-210s because of the
    /// bounded grace period and force-kill safety net.
    /// This does NOT wait for that; it returns as soon as the session is removed from the
    /// registry, so the MCP client is never blocked. The host still guarantees the PowerPoint
    /// process is fully cleaned up before it exits (see
    /// <see cref="PresentationSessionRegistry.DisposeAll()"/>).
    /// </remarks>
    private static string HandleClose(string? sessionId, bool save, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return PowerPointToolsBase.ValidationError("presentation_session_id is required for action=close.");
        }

        if (save)
        {
            if (!registry.TryGet(sessionId, out var batch))
            {
                return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
            }

            var saveResult = Commands.Save(batch);
            if (!saveResult.Success)
            {
                return SerializeResult(saveResult);
            }
        }

        var closed = registry.Close(sessionId);
        if (!closed)
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        return PowerPointToolsBase.Serialize(new
        {
            success = true,
            presentation_session_id = sessionId,
            closed = true,
            message = "Session closed; PowerPoint is shutting down in the background."
        });
    }

    private static string HandleTest(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return PowerPointToolsBase.ValidationError("filePath is required for action=test.");
        }

        return SerializeResult(Commands.Open(filePath));
    }

    private static string HandleSaveAs(
        string? sessionId,
        string? targetPath,
        PresentationSaveFormat? format,
        bool overwrite,
        PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(targetPath))
        {
            return PowerPointToolsBase.ValidationError("targetPath is required for action=save-as.");
        }

        return SerializeResult(Commands.SaveAs(
            batch,
            targetPath,
            format ?? PresentationSaveFormat.Auto,
            overwrite));
    }

    private static string HandleSaveCopyAs(
        string? sessionId,
        string? targetPath,
        bool overwrite,
        PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(targetPath))
        {
            return PowerPointToolsBase.ValidationError("targetPath is required for action=save-copy-as.");
        }

        return SerializeResult(Commands.SaveCopyAs(batch, targetPath, overwrite));
    }

    /// <summary>
    /// Lists all currently open presentation sessions.
    /// </summary>
    private static string HandleList(PresentationSessionRegistry registry)
    {
        var sessions = registry.List()
            .Select(s => new
            {
                presentation_session_id = s.SessionId,
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        return SerializeResult(Commands.GetThemeName(batch));
    }

    private static string HandleGetFinal(string? sessionId, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        return SerializeResult(Commands.GetFinal(batch));
    }

    private static string HandleSetFinal(string? sessionId, bool? isFinal, PresentationSessionRegistry registry)
    {
        if (!isFinal.HasValue)
        {
            return PowerPointToolsBase.ValidationError("isFinal is required for action=set-final.");
        }

        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        return SerializeResult(Commands.SetFinal(batch, isFinal.Value));
    }

    /// <summary>
    /// Sets a built-in document metadata property (Title, Subject, Author, Keywords, Comments,
    /// Category, Manager, or Company) on the open presentation.
    /// </summary>
    private static string HandleSetDocumentProperty(string? sessionId, string? propertyName, string? value, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
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
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return PowerPointToolsBase.ValidationError("propertyName is required for action=remove-custom-property.");
        }

        return SerializeResult(Commands.RemoveCustomProperty(batch, propertyName));
    }

    private static string HandleSetTag(
        string? sessionId,
        string? tagName,
        string? tagValue,
        PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(tagName))
        {
            return PowerPointToolsBase.ValidationError("tagName is required for action=set-tag.");
        }

        if (tagValue is null)
        {
            return PowerPointToolsBase.ValidationError("tagValue is required for action=set-tag.");
        }

        return SerializeResult(Commands.SetTag(batch, tagName, tagValue));
    }

    private static string HandleGetTag(string? sessionId, string? tagName, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(tagName))
        {
            return PowerPointToolsBase.ValidationError("tagName is required for action=get-tag.");
        }

        return SerializeResult(Commands.GetTag(batch, tagName));
    }

    private static string HandleListTags(string? sessionId, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        return SerializeResult(Commands.ListTags(batch));
    }

    private static string HandleDeleteTag(string? sessionId, string? tagName, PresentationSessionRegistry registry)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || !registry.TryGet(sessionId, out var batch))
        {
            return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {sessionId}");
        }

        if (string.IsNullOrWhiteSpace(tagName))
        {
            return PowerPointToolsBase.ValidationError("tagName is required for action=delete-tag.");
        }

        return SerializeResult(Commands.DeleteTag(batch, tagName));
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
                isFinal = result.IsFinal,
                propertyName = result.PropertyName,
                propertyValue = result.PropertyValue,
                tagName = result.TagName,
                tagValue = result.TagValue,
                tagIndex = result.TagIndex,
                tagCount = result.TagCount,
                tags = result.Tags
            });
        }

        return PowerPointToolsBase.Serialize(new
        {
            success = false,
            errorMessage = result.ErrorMessage,
            presentationPath = result.PresentationPath,
            themeName = result.ThemeName,
            isFinal = result.IsFinal,
            propertyName = result.PropertyName,
            propertyValue = result.PropertyValue,
            tagName = result.TagName,
            tagValue = result.TagValue,
            tagIndex = result.TagIndex,
            tagCount = result.TagCount,
            tags = result.Tags,
            isError = true
        });
    }
}
