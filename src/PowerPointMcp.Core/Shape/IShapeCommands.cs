using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Shape;

/// <summary>
/// Shape commands: add rectangles/text boxes, count, delete, reposition/resize.
/// Operates within an already-open IPresentationBatch, targeting a specific slide by
/// its 1-based index.
/// </summary>
[ServiceCategory("shape", "Shape")]
[McpTool("shape", Title = "Shape Operations", Destructive = true, Category = "content",
    Description = "Add, count, delete, reposition, and resize shapes on a slide in an open presentation session.")]
public interface IShapeCommands
{
    /// <summary>Adds a rectangle shape to the given slide.</summary>
    ShapeOperationResult AddRectangle(ComInterop.Session.IPresentationBatch batch, int slideIndex, float left, float top, float width, float height);

    /// <summary>Adds a text box with the given text to the given slide.</summary>
    ShapeOperationResult AddTextBox(ComInterop.Session.IPresentationBatch batch, int slideIndex, float left, float top, float width, float height, string text);

    /// <summary>
    /// Adds a non-rectangle "auto shape" (oval, diamond, arrow, star bracket, etc.) to the given
    /// slide, identified by its <c>MsoAutoShapeType</c> enum member name (e.g.
    /// <c>"msoShapeOval"</c>, <c>"msoShapeRightArrow"</c>). See <c>slides-and-shapes.md</c> for
    /// the full supported name list.
    /// </summary>
    ShapeOperationResult AddAutoShape(ComInterop.Session.IPresentationBatch batch, int slideIndex, string shapeType, float left, float top, float width, float height);

    /// <summary>
    /// Adds a straight line from (<paramref name="beginX"/>, <paramref name="beginY"/>) to
    /// (<paramref name="endX"/>, <paramref name="endY"/>) on the given slide.
    /// </summary>
    ShapeOperationResult AddLine(ComInterop.Session.IPresentationBatch batch, int slideIndex, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// Adds a connector shape (straight, elbow, or curved) between two points on the given slide,
    /// identified by its <c>MsoConnectorType</c> enum member name (<c>"msoConnectorStraight"</c>,
    /// <c>"msoConnectorElbow"</c>, or <c>"msoConnectorCurve"</c>). Unlike
    /// <see cref="AddLine"/>, a connector is intended to visually link two shapes, but this
    /// command creates it as a free-floating shape between raw coordinates — attaching it to
    /// specific shapes is not exposed by this tool surface.
    /// </summary>
    ShapeOperationResult AddConnector(ComInterop.Session.IPresentationBatch batch, int slideIndex, string connectorType, float beginX, float beginY, float endX, float endY);

    /// <summary>Gets the number of shapes on the given slide.</summary>
    ShapeOperationResult GetCount(ComInterop.Session.IPresentationBatch batch, int slideIndex);

    /// <summary>Deletes the shape at the given 1-based index on the given slide.</summary>
    ShapeOperationResult Delete(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the position of a shape on the given slide.</summary>
    ShapeOperationResult SetPosition(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, float left, float top);

    /// <summary>Sets the size of a shape on the given slide.</summary>
    ShapeOperationResult SetSize(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, float width, float height);

    /// <summary>Sets a shape's fill to a solid RGB color.</summary>
    ShapeOperationResult SetFill(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue);

    /// <summary>Gets a shape's solid fill color.</summary>
    ShapeOperationResult GetFill(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets one or more line/border properties of a shape. Any parameter left null is unchanged.
    /// Passing <paramref name="red"/>/<paramref name="green"/>/<paramref name="blue"/> together
    /// sets the line color; <paramref name="dashStyle"/> is an <c>MsoLineDashStyle</c> enum
    /// member name (e.g. <c>"msoLineSolid"</c>, <c>"msoLineDash"</c>).
    /// </summary>
    ShapeOperationResult SetLine(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        byte? red = null,
        byte? green = null,
        byte? blue = null,
        float? weight = null,
        string? dashStyle = null,
        bool? visible = null);

    /// <summary>Gets a shape's line/border color, weight, dash style, and visibility.</summary>
    ShapeOperationResult GetLine(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets a shape's rotation, in degrees clockwise from its upright position.</summary>
    ShapeOperationResult SetRotation(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, float degrees);

    /// <summary>Gets a shape's rotation, in degrees clockwise from its upright position.</summary>
    ShapeOperationResult GetRotation(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Flips a shape horizontally or vertically in place (<paramref name="direction"/>: <c>"horizontal"</c> or <c>"vertical"</c>).</summary>
    ShapeOperationResult Flip(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string direction);

    /// <summary>
    /// Moves a shape's position in the slide's z-order (draw order). <paramref name="zOrderCommand"/>
    /// is one of <c>"bring-to-front"</c>, <c>"send-to-back"</c>, <c>"bring-forward"</c>, or
    /// <c>"send-backward"</c>.
    /// </summary>
    ShapeOperationResult SetZOrder(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string zOrderCommand);

    /// <summary>Turns a shape's default drop shadow on or off.</summary>
    ShapeOperationResult SetShadow(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible);

    /// <summary>Gets whether a shape's drop shadow is visible.</summary>
    ShapeOperationResult GetShadow(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Groups two or more shapes on the given slide into a single shape, identified by their
    /// 1-based shape indices. Returns the new grouped shape's index.
    /// </summary>
    ShapeOperationResult Group(ComInterop.Session.IPresentationBatch batch, int slideIndex, IReadOnlyList<int> shapeIndexes);

    /// <summary>Ungroups a previously-grouped shape back into its individual member shapes.</summary>
    ShapeOperationResult Ungroup(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets a shape's name (as shown in the Selection Pane).</summary>
    ShapeOperationResult SetName(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string name);

    /// <summary>Gets a shape's name.</summary>
    ShapeOperationResult GetName(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets a shape's alternative text (accessibility description).</summary>
    ShapeOperationResult SetAltText(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string altText);

    /// <summary>Gets a shape's alternative text (accessibility description).</summary>
    ShapeOperationResult GetAltText(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets a shape's mouse-click hyperlink to <paramref name="address"/> (a URL, e.g.
    /// <c>"https://example.com"</c>, or a local file path). Optionally sets the hyperlink's
    /// screen tip text shown on hover. Clicking the shape at presentation time navigates to
    /// <paramref name="address"/>.
    /// </summary>
    ShapeOperationResult SetHyperlink(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string address, string? screenTip = null);

    /// <summary>
    /// Gets a shape's mouse-click hyperlink, if any. Returns <c>HasHyperlink = false</c> (with
    /// null <c>HyperlinkAddress</c>/<c>HyperlinkScreenTip</c>) when the shape has no hyperlink
    /// action assigned.
    /// </summary>
    ShapeOperationResult GetHyperlink(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Removes a shape's mouse-click hyperlink, if any (idempotent — no-op if none is set).</summary>
    ShapeOperationResult RemoveHyperlink(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);
}
