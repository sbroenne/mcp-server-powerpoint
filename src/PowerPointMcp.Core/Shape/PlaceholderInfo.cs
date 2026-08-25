namespace Sbroenne.PowerPointMcp.Core.Shape;

/// <summary>State and geometry of one native PowerPoint placeholder.</summary>
public sealed class PlaceholderInfo
{
    /// <summary>1-based index in the slide's Shapes collection.</summary>
    public required int ShapeIndex { get; init; }

    /// <summary>Native PpPlaceholderType member name.</summary>
    public required string PlaceholderType { get; init; }

    /// <summary>Shape name shown in PowerPoint's Selection Pane.</summary>
    public required string Name { get; init; }

    /// <summary>Alternative text assigned to the placeholder.</summary>
    public required string AltText { get; init; }

    /// <summary>Horizontal position in points.</summary>
    public required float Left { get; init; }

    /// <summary>Vertical position in points.</summary>
    public required float Top { get; init; }

    /// <summary>Width in points.</summary>
    public required float Width { get; init; }

    /// <summary>Height in points.</summary>
    public required float Height { get; init; }

    /// <summary>Whether the placeholder currently contains non-empty text.</summary>
    public required bool HasText { get; init; }

    /// <summary>Current text when the placeholder contains text.</summary>
    public string? Text { get; init; }

    /// <summary>Whether the placeholder currently contains an image fill.</summary>
    public required bool HasImage { get; init; }

    /// <summary>Stable content state: empty, text, or image.</summary>
    public required string ContentState { get; init; }
}
