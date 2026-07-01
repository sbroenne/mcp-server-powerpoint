using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <summary>
/// Slide layout commands: apply/read a slide's built-in layout. Operates within an
/// already-open IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
public interface ILayoutCommands
{
    /// <summary>
    /// Applies a built-in slide layout by its <c>PpSlideLayout</c> enum member name
    /// (e.g. "ppLayoutBlank", "ppLayoutTitleOnly", "ppLayoutText").
    /// </summary>
    LayoutOperationResult SetLayout(IPresentationBatch batch, int slideIndex, string layoutName);

    /// <summary>Gets the current slide's layout name.</summary>
    LayoutOperationResult GetLayout(IPresentationBatch batch, int slideIndex);
}
