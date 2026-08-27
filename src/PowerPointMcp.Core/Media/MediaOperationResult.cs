namespace Sbroenne.PowerPointMcp.Core.Media;

/// <summary>Result of a media operation.</summary>
public sealed class MediaOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the media shape.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>Total shape count on the slide.</summary>
    public int? ShapeCount { get; init; }

    /// <summary>Typed <c>PpMediaType</c> member name reported by PowerPoint.</summary>
    public string? MediaTypeName { get; init; }

    /// <summary>Full source path used for insertion.</summary>
    public string? SourcePath { get; init; }

    /// <summary>Whether insertion linked the media shape to its source file.</summary>
    public bool? LinkToFile { get; init; }

    /// <summary>Whether insertion saved media data with the presentation.</summary>
    public bool? SaveWithDocument { get; init; }

    /// <summary>Horizontal position in points.</summary>
    public float? Left { get; init; }

    /// <summary>Vertical position in points.</summary>
    public float? Top { get; init; }

    /// <summary>Width in points.</summary>
    public float? Width { get; init; }

    /// <summary>Height in points.</summary>
    public float? Height { get; init; }
}
