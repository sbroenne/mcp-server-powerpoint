namespace Sbroenne.PowerPointMcp.Core.CustomShow;

/// <summary>
/// Result of a named-custom-show operation (list, create, delete).
/// </summary>
public sealed class CustomShowOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false; null/empty when true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Name of the custom show acted upon, for Create/Delete.</summary>
    public string? Name { get; init; }

    /// <summary>
    /// 1-based current slide indices in playback order, for Create. Reflects the slide positions
    /// at creation time; slides may later move without updating this snapshot.
    /// </summary>
    public IReadOnlyList<int>? SlideIndices { get; init; }

    /// <summary>All custom shows in the presentation, in collection order, for List.</summary>
    public IReadOnlyList<CustomShowEntry>? Shows { get; init; }

    /// <summary>One named custom show entry within a List result.</summary>
    public sealed class CustomShowEntry
    {
        /// <summary>1-based position of this custom show within the presentation's collection.</summary>
        public int Index { get; init; }

        /// <summary>Name of the custom show.</summary>
        public string Name { get; init; } = string.Empty;

        /// <summary>
        /// 1-based current slide indices in playback order. A slide that has since been deleted is
        /// omitted, so this list can be shorter than the show's original slide count.
        /// </summary>
        public IReadOnlyList<int> SlideIndices { get; init; } = Array.Empty<int>();
    }
}
