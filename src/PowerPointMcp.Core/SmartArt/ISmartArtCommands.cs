using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.SmartArt;

/// <summary>
/// SmartArt commands: add a SmartArt diagram to a slide from PowerPoint's built-in layout
/// gallery, and add/read/update/delete/count the diagram's nodes. Operates within an
/// already-open <see cref="ComInterop.Session.IPresentationBatch"/>, targeting a specific slide
/// and shape by 1-based index.
/// </summary>
/// <remarks>
/// Layouts are identified by their gallery display name (e.g. <c>"Basic Process"</c>,
/// <c>"Organization Chart"</c>, <c>"Basic Cycle"</c>) rather than a numeric gallery position or
/// file path — verified live against <c>Application.SmartArtLayouts</c> (176 layouts on a
/// standard install), where <c>.Name</c> is a stable, human-readable identifier and gallery order
/// is not. See <c>smart-art.md</c> for the full supported name list.
/// <para>
/// Nodes are addressed by their 1-based position in the diagram's flat <c>SmartArt.AllNodes</c>
/// collection (document order across the whole hierarchy), matching how PowerPoint's own object
/// model already exposes them — no invented addressing scheme. <see cref="AddNode"/> appends a
/// new top-level (root) node; <see cref="AddChildNode"/> appends a new node as the last child of
/// an existing node, addressed the same way.
/// </para>
/// </remarks>
[ServiceCategory("smartart", "SmartArt")]
[McpTool("smartart", Title = "SmartArt Operations", Destructive = true, Category = "content",
    Description = "Add a SmartArt diagram to a slide and add, read, update, delete, or count its nodes in an open presentation session.")]
public interface ISmartArtCommands
{
    /// <summary>
    /// Adds a SmartArt diagram to the given slide using the built-in gallery layout identified by
    /// <paramref name="layoutName"/> (e.g. <c>"Basic Process"</c>, <c>"Organization Chart"</c>).
    /// The diagram is created with whatever default nodes that layout starts with (commonly 3
    /// placeholder nodes for list/process layouts, 1 root node for hierarchy layouts).
    /// </summary>
    SmartArtOperationResult AddSmartArt(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        string layoutName,
        float left,
        float top,
        float width,
        float height);

    /// <summary>
    /// Adds a new top-level (root) node to the SmartArt diagram at <paramref name="shapeIndex"/>
    /// on the given slide, with the given text, appended after the existing top-level nodes.
    /// </summary>
    SmartArtOperationResult AddNode(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string text);

    /// <summary>
    /// Adds a new child node to the SmartArt diagram at <paramref name="shapeIndex"/> on the
    /// given slide, nested under the node at <paramref name="parentNodeIndex"/> (1-based
    /// position in <c>SmartArt.AllNodes</c>), appended after that parent's existing children.
    /// </summary>
    SmartArtOperationResult AddChildNode(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int parentNodeIndex,
        string text);

    /// <summary>
    /// Sets the text of the node at <paramref name="nodeIndex"/> (1-based position in
    /// <c>SmartArt.AllNodes</c>) in the SmartArt diagram at <paramref name="shapeIndex"/> on the
    /// given slide.
    /// </summary>
    SmartArtOperationResult SetNodeText(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int nodeIndex,
        string text);

    /// <summary>
    /// Gets the text of the node at <paramref name="nodeIndex"/> (1-based position in
    /// <c>SmartArt.AllNodes</c>) in the SmartArt diagram at <paramref name="shapeIndex"/> on the
    /// given slide.
    /// </summary>
    SmartArtOperationResult GetNodeText(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int nodeIndex);

    /// <summary>
    /// Deletes the node at <paramref name="nodeIndex"/> (1-based position in
    /// <c>SmartArt.AllNodes</c>), and any of its children, from the SmartArt diagram at
    /// <paramref name="shapeIndex"/> on the given slide.
    /// </summary>
    SmartArtOperationResult DeleteNode(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int nodeIndex);

    /// <summary>
    /// Gets the total number of nodes (<c>SmartArt.AllNodes.Count</c>, flat across the whole
    /// hierarchy) in the SmartArt diagram at <paramref name="shapeIndex"/> on the given slide.
    /// </summary>
    SmartArtOperationResult GetNodeCount(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex);
}
