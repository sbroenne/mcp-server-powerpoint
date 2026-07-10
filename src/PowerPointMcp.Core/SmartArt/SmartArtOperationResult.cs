namespace Sbroenne.PowerPointMcp.Core.SmartArt;

/// <summary>
/// Result of a SmartArt operation (add diagram, add/read/update/delete node, node count).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1):
/// Success == true implies ErrorMessage is null/empty.
/// </remarks>
public sealed class SmartArtOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the SmartArt shape the operation created or acted on, if applicable.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Total shape count on the slide after the operation, for AddSmartArt.</summary>
    public int? ShapeCount { get; init; }

    /// <summary>The gallery layout name used to create the diagram, for AddSmartArt.</summary>
    public string? LayoutName { get; init; }

    /// <summary>
    /// 1-based position (in <c>SmartArt.AllNodes</c>) of the node the operation created or
    /// acted on, if applicable.
    /// </summary>
    public int? NodeIndex { get; init; }

    /// <summary>Total node count in the diagram (<c>SmartArt.AllNodes.Count</c>) after the operation, if applicable.</summary>
    public int? NodeCount { get; init; }

    /// <summary>Node text, for GetNodeText/SetNodeText/AddNode/AddChildNode.</summary>
    public string? NodeText { get; init; }
}
