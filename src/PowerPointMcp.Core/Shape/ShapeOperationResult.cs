namespace Sbroenne.PowerPointMcp.Core.Shape;

/// <summary>
/// Result of a shape operation (add, delete, count, position/size).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1):
/// Success == true implies ErrorMessage is null/empty.
/// </remarks>
public sealed class ShapeOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the shape the operation created or acted on, if applicable.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Total shape count on the slide after the operation.</summary>
    public int? ShapeCount { get; init; }

    /// <summary>Shape left position in points, if applicable.</summary>
    public float? Left { get; init; }

    /// <summary>Shape top position in points, if applicable.</summary>
    public float? Top { get; init; }

    /// <summary>Shape width in points, if applicable.</summary>
    public float? Width { get; init; }

    /// <summary>Shape height in points, if applicable.</summary>
    public float? Height { get; init; }
}
