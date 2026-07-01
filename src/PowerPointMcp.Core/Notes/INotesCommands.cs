using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Notes;

/// <summary>
/// Speaker notes commands: set/get the notes text for a slide. Operates within an
/// already-open IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
public interface INotesCommands
{
    /// <summary>Sets the speaker notes text for a slide.</summary>
    NotesOperationResult SetNotesText(IPresentationBatch batch, int slideIndex, string text);

    /// <summary>Gets the speaker notes text for a slide.</summary>
    NotesOperationResult GetNotesText(IPresentationBatch batch, int slideIndex);
}
