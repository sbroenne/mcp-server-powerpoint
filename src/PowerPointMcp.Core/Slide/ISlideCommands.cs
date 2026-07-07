using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>
/// Slide lifecycle commands: add, delete, count. First domain built on top of the
/// presentation lifecycle commands, operating within an already-open IPresentationBatch.
/// </summary>
[ServiceCategory("slide", "Slide")]
[McpTool("slide", Title = "Slide Operations", Destructive = true, Category = "content",
    Description = "Add, count, and delete slides in an open presentation session.")]
public interface ISlideCommands
{
    /// <summary>
    /// Adds a new blank slide at the end of the presentation.
    /// </summary>
    /// <returns>A result with Success, the new slide's 1-based index, and the new total count.</returns>
    SlideOperationResult AddBlank(IPresentationBatch batch);

    /// <summary>
    /// Gets the current number of slides in the presentation.
    /// </summary>
    SlideOperationResult GetCount(IPresentationBatch batch);

    /// <summary>
    /// Deletes the slide at the given 1-based index.
    /// </summary>
    SlideOperationResult Delete(IPresentationBatch batch, int slideIndex);
}
