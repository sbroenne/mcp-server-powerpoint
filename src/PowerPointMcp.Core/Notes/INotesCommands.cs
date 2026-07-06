using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Notes;

/// <summary>
/// Speaker notes commands: set/get the notes text for a slide. Operates within an
/// already-open IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
[ServiceCategory("notes", "Notes")]
[McpTool("notes", Title = "Speaker Notes Operations", Destructive = true, Category = "content",
    Description = "Set or get the speaker notes text for a slide in an open presentation session.")]
public interface INotesCommands
{
    /// <summary>Sets the speaker notes text for a slide.</summary>
    NotesOperationResult SetNotesText(IPresentationBatch batch, int slideIndex, string text);

    /// <summary>Gets the speaker notes text for a slide.</summary>
    NotesOperationResult GetNotesText(IPresentationBatch batch, int slideIndex);
}
