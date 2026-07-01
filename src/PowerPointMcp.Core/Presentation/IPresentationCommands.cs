using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>
/// Presentation lifecycle commands: create, close, save.
/// </summary>
/// <remarks>
/// First Core command domain implemented, proving the ComInterop batch pattern end-to-end.
/// Remaining domains from the plan (Slide, Shape, TextFrame, Table, Chart, Image, Notes,
/// Layout/Master, Export/QA) are follow-up work — see plan.md continuation notes.
/// </remarks>
public interface IPresentationCommands
{
    /// <summary>
    /// Creates a new, empty presentation at the given path and saves it immediately.
    /// </summary>
    PresentationOperationResult Create(string filePath, bool isMacroEnabled = false);

    /// <summary>
    /// Saves the presentation currently open in the given batch.
    /// </summary>
    PresentationOperationResult Save(IPresentationBatch batch);
}
