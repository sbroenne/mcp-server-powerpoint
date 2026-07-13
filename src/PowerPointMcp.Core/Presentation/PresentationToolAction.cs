using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>
/// Actions dispatched by the single hand-written MCP <c>presentation</c> tool (and matched by
/// the CLI's <c>pptcli presentation &lt;action&gt;</c> command). Mirrors mcp-server-excel's
/// <c>FileAction</c> enum shape — one action-dispatch tool per domain, including session
/// lifecycle, instead of one MCP tool per verb.
/// </summary>
/// <remarks>
/// Session lifecycle (<see cref="Create"/>, <see cref="Open"/>, <see cref="Save"/>,
/// <see cref="Close"/>, <see cref="List"/>) is handled directly against
/// <see cref="Sbroenne.PowerPointMcp.ComInterop.Session.PresentationSessionRegistry"/> by the MCP
/// tool, NOT via <see cref="IPresentationCommands.Open"/>/<see cref="IPresentationCommands.Create"/>
/// (those Core methods are standalone file-validation utilities with different semantics — see
/// their XML docs — and are exposed separately via the generated CLI <c>presentation</c>
/// command). The remaining actions map 1:1 to <see cref="IPresentationCommands"/> methods.
/// </remarks>
public enum PresentationToolAction
{
    /// <summary>Create a new, empty presentation file on disk and leave a session open.</summary>
    [JsonStringEnumMemberName("create")]
    Create,

    /// <summary>Open an existing presentation file and start a session.</summary>
    [JsonStringEnumMemberName("open")]
    Open,

    /// <summary>Save the presentation for an open session to its current file.</summary>
    [JsonStringEnumMemberName("save")]
    Save,

    /// <summary>Close an open session and release its PowerPoint process.</summary>
    [JsonStringEnumMemberName("close")]
    Close,

    /// <summary>List all currently open presentation sessions.</summary>
    [JsonStringEnumMemberName("list")]
    List,

    /// <summary>Apply a template's masters/theme/layouts to the open presentation.</summary>
    [JsonStringEnumMemberName("apply-template")]
    ApplyTemplate,

    /// <summary>Get the design/theme name currently applied to the presentation.</summary>
    [JsonStringEnumMemberName("get-theme-name")]
    GetThemeName,

    /// <summary>Set a built-in document metadata property (Title, Author, etc.).</summary>
    [JsonStringEnumMemberName("set-document-property")]
    SetDocumentProperty,

    /// <summary>Get a built-in document metadata property (Title, Author, etc.).</summary>
    [JsonStringEnumMemberName("get-document-property")]
    GetDocumentProperty,

    /// <summary>Create or update a custom (user-defined) document property.</summary>
    [JsonStringEnumMemberName("set-custom-property")]
    SetCustomProperty,

    /// <summary>Get a custom (user-defined) document property.</summary>
    [JsonStringEnumMemberName("get-custom-property")]
    GetCustomProperty,

    /// <summary>Remove a custom (user-defined) document property.</summary>
    [JsonStringEnumMemberName("remove-custom-property")]
    RemoveCustomProperty
}

/// <summary>
/// Maps <see cref="PresentationToolAction"/> values to their kebab-case wire names, mirroring
/// mcp-server-excel's <c>ActionExtensions.ToActionString</c> pattern.
/// </summary>
public static class PresentationToolActionExtensions
{
    /// <summary>Maps a <see cref="PresentationToolAction"/> value to its kebab-case wire name.</summary>
    public static string ToActionString(this PresentationToolAction action) => action switch
    {
        PresentationToolAction.Create => "create",
        PresentationToolAction.Open => "open",
        PresentationToolAction.Save => "save",
        PresentationToolAction.Close => "close",
        PresentationToolAction.List => "list",
        PresentationToolAction.ApplyTemplate => "apply-template",
        PresentationToolAction.GetThemeName => "get-theme-name",
        PresentationToolAction.SetDocumentProperty => "set-document-property",
        PresentationToolAction.GetDocumentProperty => "get-document-property",
        PresentationToolAction.SetCustomProperty => "set-custom-property",
        PresentationToolAction.GetCustomProperty => "get-custom-property",
        PresentationToolAction.RemoveCustomProperty => "remove-custom-property",
        _ => throw new ArgumentOutOfRangeException(nameof(action), action, "Unknown PresentationToolAction")
    };
}
