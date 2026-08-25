namespace Sbroenne.PowerPointMcp.Core.Accessibility;

/// <summary>Result of an accessibility operation.</summary>
public sealed class AccessibilityOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Deterministic accessibility issues found by an audit.</summary>
    public IReadOnlyList<AccessibilityIssue>? Issues { get; init; }

    /// <summary>Shape reading order expressed as 1-based shape indexes.</summary>
    public IReadOnlyList<int>? ReadingOrder { get; init; }

    /// <summary>One deterministic accessibility issue.</summary>
    public sealed class AccessibilityIssue
    {
        /// <summary>Stable machine-readable issue code.</summary>
        public required string Code { get; init; }

        /// <summary>Human-readable issue description.</summary>
        public required string Message { get; init; }

        /// <summary>1-based slide index.</summary>
        public int SlideIndex { get; init; }

        /// <summary>1-based shape index, when the issue targets a shape.</summary>
        public int? ShapeIndex { get; init; }
    }
}
