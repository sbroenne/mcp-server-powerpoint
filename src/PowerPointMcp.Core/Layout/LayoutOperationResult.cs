namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <summary>
/// Result of a slide layout operation (set/get layout).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class LayoutOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// The slide layout name (a <c>Microsoft.Office.Interop.PowerPoint.PpSlideLayout</c>
    /// enum member name, e.g. "ppLayoutBlank", "ppLayoutTitleOnly", "ppLayoutText").
    /// </summary>
    public string? LayoutName { get; init; }

    /// <summary>1-based index of the slide master for ListLayouts.</summary>
    public int? MasterIndex { get; init; }

    /// <summary>Name of the slide master for ListLayouts.</summary>
    public string? MasterName { get; init; }

    /// <summary>Layouts attached to the selected slide master.</summary>
    public IReadOnlyList<LayoutInventoryEntry>? Layouts { get; init; }

    /// <summary>Represents one layout entry inside a slide master's inventory.</summary>
    public sealed class LayoutInventoryEntry
    {
        /// <summary>1-based index of the layout within the slide master.</summary>
        public int LayoutIndex { get; init; }

        /// <summary>Name of the layout.</summary>
        public string? LayoutName { get; init; }

        /// <summary>Whether any slide in the presentation currently uses this layout.</summary>
        public bool IsUsed { get; init; }
    }
}
