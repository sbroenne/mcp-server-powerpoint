using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <summary>
/// Slide layout commands: apply/read a slide's built-in layout. Operates within an
/// already-open IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
[ServiceCategory("layout", "Layout")]
[McpTool("layout", Title = "Slide Layout Operations", Destructive = true, Category = "content",
    Description = "Apply or read a slide's built-in layout in an open presentation session.")]
public interface ILayoutCommands
{
    /// <summary>
    /// Applies a built-in slide layout by its <c>PpSlideLayout</c> enum member name
    /// (e.g. "ppLayoutBlank", "ppLayoutTitleOnly", "ppLayoutText").
    /// </summary>
    LayoutOperationResult SetLayout(IPresentationBatch batch, int slideIndex, string layoutName);

    /// <summary>Gets the current slide's layout name.</summary>
    LayoutOperationResult GetLayout(IPresentationBatch batch, int slideIndex);

    /// <summary>Lists the layouts attached to a given slide master, including whether any slide uses each one.</summary>
    LayoutOperationResult ListLayouts(IPresentationBatch batch, int masterIndex);

    /// <summary>Deletes an unused layout from a slide master. Fails if any slide still references it.</summary>
    LayoutOperationResult DeleteLayout(IPresentationBatch batch, int masterIndex, int layoutIndex);
}
