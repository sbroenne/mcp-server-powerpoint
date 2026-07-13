using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>
/// Presentation lifecycle commands: create, close, save.
/// </summary>
/// <remarks>
/// First Core command domain implemented, proving the ComInterop batch pattern end-to-end.
/// Remaining domains from the plan (Slide, Shape, TextFrame, Table, Chart, Image, Notes,
/// Layout/Master, Export/QA) are follow-up work — see plan.md continuation notes.
///
/// Deliberately carries NO <c>[ServiceCategory]</c>/<c>[McpTool]</c> attribute, unlike every
/// other Core command domain — this mirrors mcp-server-excel's <c>IFileCommands</c> exactly.
/// Session-establishing operations (<see cref="Create"/>, <see cref="Open"/>) can't fit the
/// generic generators' fixed, non-nullable <c>session_id</c>/action-dispatch shape, so this
/// whole domain is exposed via hand-written surfaces instead: the MCP <c>presentation</c> tool
/// (<c>PresentationTools.cs</c>) and the CLI's hand-written <c>session</c> branch
/// (<c>SessionCommands.cs</c> + <c>PowerPointMcpService.HandleSessionCommand</c>) — zero
/// generator involvement, zero duplication, matching Excel's architecture 1:1.
/// </remarks>
public interface IPresentationCommands
{
    /// <summary>
    /// Creates a new, empty presentation at the given path and saves it immediately.
    /// </summary>
    PresentationOperationResult Create(string filePath, bool isMacroEnabled = false);

    /// <summary>
    /// Validates that a presentation file exists and can be opened by PowerPoint, then closes
    /// it again. Rule 1/1b: a missing file is expected/graceful input and returns
    /// <c>Success = false</c> with an <see cref="PresentationOperationResult.ErrorMessage"/> —
    /// it does NOT start PowerPoint. Once the file is confirmed to exist, this opens (and
    /// immediately closes) a real batch to prove PowerPoint can actually open it; unexpected COM
    /// failures during that step (corrupt file, PowerPoint not installed, etc.) propagate.
    /// </summary>
    /// <remarks>
    /// This does NOT keep a session open for further edits — callers that want to keep editing
    /// must call <see cref="Sbroenne.PowerPointMcp.ComInterop.Session.PresentationSession.BeginBatch(string)"/>
    /// themselves (this mirrors <see cref="Create"/>, which also creates+saves+closes rather than
    /// returning a live session). "Closing" a presentation is simply disposing the
    /// <see cref="Sbroenne.PowerPointMcp.ComInterop.Session.IPresentationBatch"/> obtained from
    /// <c>BeginBatch</c>/<c>CreateNew</c> — there is no separate Core-level Close() command
    /// because <c>IPresentationBatch.Dispose()</c> already IS the close operation (it drives
    /// <see cref="Sbroenne.PowerPointMcp.ComInterop.Session.PresentationShutdownService"/>'s
    /// resilient close/quit). Adding a Close() wrapper here would just re-expose Dispose() under
    /// a different name.
    /// </remarks>
    PresentationOperationResult Open(string filePath);

    /// <summary>
    /// Saves the presentation currently open in the given batch.
    /// </summary>
    PresentationOperationResult Save(IPresentationBatch batch);

    /// <summary>
    /// Applies a PowerPoint template's masters/theme/layouts to the presentation currently open
    /// in the given batch, preserving all existing slide content. Wraps the COM API
    /// <c>Presentation.ApplyTemplate(templatePath)</c>.
    /// </summary>
    /// <param name="batch">The open batch whose presentation will be restyled.</param>
    /// <param name="templatePath">Full path to a <c>.potx</c>/<c>.potm</c>/<c>.pot</c> template file (a <c>.pptx</c>/<c>.pptm</c> presentation may also be used as a template source, matching PowerPoint's own behavior).</param>
    PresentationOperationResult ApplyTemplate(IPresentationBatch batch, string templatePath);

    /// <summary>
    /// Reads the name of the design/theme currently applied to the presentation open in the
    /// given batch — useful for verifying that <see cref="ApplyTemplate"/> actually changed the
    /// presentation's styling.
    /// </summary>
    PresentationOperationResult GetThemeName(IPresentationBatch batch);

    /// <summary>
    /// Sets a built-in document metadata property (Title, Subject, Author, Keywords, Comments,
    /// Category, Manager, or Company) on the presentation open in the given batch. Wraps
    /// <c>Presentation.BuiltInDocumentProperties[name].Value</c>.
    /// </summary>
    /// <param name="batch">The open batch whose presentation metadata will be updated.</param>
    /// <param name="propertyName">One of the supported built-in property names (case-insensitive).</param>
    /// <param name="value">The new value for the property.</param>
    PresentationOperationResult SetDocumentProperty(IPresentationBatch batch, string propertyName, string value);

    /// <summary>
    /// Reads a built-in document metadata property (Title, Subject, Author, Keywords, Comments,
    /// Category, Manager, or Company) from the presentation open in the given batch.
    /// </summary>
    /// <param name="batch">The open batch whose presentation metadata will be read.</param>
    /// <param name="propertyName">One of the supported built-in property names (case-insensitive).</param>
    PresentationOperationResult GetDocumentProperty(IPresentationBatch batch, string propertyName);

    /// <summary>
    /// Creates or updates a custom (user-defined) string document property on the presentation
    /// open in the given batch. Wraps <c>Presentation.CustomDocumentProperties</c>.
    /// </summary>
    /// <param name="batch">The open batch whose presentation metadata will be updated.</param>
    /// <param name="propertyName">The custom property's name.</param>
    /// <param name="value">The custom property's string value.</param>
    PresentationOperationResult SetCustomProperty(IPresentationBatch batch, string propertyName, string value);

    /// <summary>
    /// Reads a custom (user-defined) document property from the presentation open in the given
    /// batch. Returns <c>Success = false</c> if no custom property with that name exists.
    /// </summary>
    /// <param name="batch">The open batch whose presentation metadata will be read.</param>
    /// <param name="propertyName">The custom property's name.</param>
    PresentationOperationResult GetCustomProperty(IPresentationBatch batch, string propertyName);

    /// <summary>
    /// Removes a custom (user-defined) document property from the presentation open in the
    /// given batch. Returns <c>Success = false</c> if no custom property with that name exists.
    /// </summary>
    /// <param name="batch">The open batch whose presentation metadata will be updated.</param>
    /// <param name="propertyName">The custom property's name.</param>
    PresentationOperationResult RemoveCustomProperty(IPresentationBatch batch, string propertyName);
}
