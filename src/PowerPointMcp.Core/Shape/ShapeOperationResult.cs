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

    /// <summary>The MsoAutoShapeType name used to create the shape, for AddAutoShape.</summary>
    public string? ShapeTypeName { get; init; }

    /// <summary>The MsoConnectorType name used to create the connector, for AddConnector.</summary>
    public string? ConnectorTypeName { get; init; }

    /// <summary>X coordinate of the line/connector's start point in points, if applicable.</summary>
    public float? BeginX { get; init; }

    /// <summary>Y coordinate of the line/connector's start point in points, if applicable.</summary>
    public float? BeginY { get; init; }

    /// <summary>X coordinate of the line/connector's end point in points, if applicable.</summary>
    public float? EndX { get; init; }

    /// <summary>Y coordinate of the line/connector's end point in points, if applicable.</summary>
    public float? EndY { get; init; }

    /// <summary>Fill or line color as an RGB integer (0xBBGGRR, PowerPoint's native color order), if applicable.</summary>
    public int? ColorRgb { get; init; }

    /// <summary>Line/border weight in points, if applicable.</summary>
    public float? LineWeight { get; init; }

    /// <summary>The MsoLineDashStyle name of the shape's line/border, if applicable.</summary>
    public string? DashStyleName { get; init; }

    /// <summary>Whether the shape's line/border or shadow is visible, if applicable.</summary>
    public bool? Visible { get; init; }

    /// <summary>Shape rotation in degrees clockwise from upright, if applicable.</summary>
    public float? Rotation { get; init; }

    /// <summary>The flip direction applied ("horizontal" or "vertical"), if applicable.</summary>
    public string? FlipDirection { get; init; }

    /// <summary>The z-order command applied, if applicable.</summary>
    public string? ZOrderCommand { get; init; }

    /// <summary>Number of shapes produced by an Ungroup operation, if applicable.</summary>
    public int? UngroupedShapeCount { get; init; }

    /// <summary>Shape name, if applicable.</summary>
    public string? Name { get; init; }

    /// <summary>Shape alternative text (accessibility description), if applicable.</summary>
    public string? AltText { get; init; }

    /// <summary>Whether the shape has a mouse-click hyperlink assigned, if applicable.</summary>
    public bool? HasHyperlink { get; init; }

    /// <summary>The shape's hyperlink target address (URL or file path), if applicable.</summary>
    public string? HyperlinkAddress { get; init; }

    /// <summary>The shape's hyperlink hover screen tip text, if applicable.</summary>
    public string? HyperlinkScreenTip { get; init; }
}
