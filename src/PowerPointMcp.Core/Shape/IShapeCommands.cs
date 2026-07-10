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

    /// <summary>
    /// Turns a shape's drop shadow on or off and, when <paramref name="visible"/> is true, sets
    /// its color/transparency/blur/offset. The color/formatting parameters are optional additions
    /// to the original visibility-only overload — existing callers passing only <paramref name="visible"/>
    /// remain valid and get PowerPoint's shadow defaults.
    /// </summary>
    ShapeOperationResult SetShadow(
        ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible,
        byte red = 0, byte green = 0, byte blue = 0,
        float transparency = 0.5f, float blur = 4f, float offsetX = 3f, float offsetY = 3f);

    /// <summary>Gets a shape's drop shadow visibility and, when visible, its color/transparency/blur/offset.</summary>
    ShapeOperationResult GetShadow(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Applies a glow effect to a shape with the given color, radius (in points), and transparency (0-1). A radius of 0 removes the glow.</summary>
    ShapeOperationResult SetGlow(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue, float radius, float transparency = 0f);

    /// <summary>Gets a shape's glow color, radius, and transparency.</summary>
    ShapeOperationResult GetGlow(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Turns a shape's reflection effect on or off and, when visible, sets its transparency, size (% of shape height), and blur.</summary>
    ShapeOperationResult SetReflection(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible, float transparency = 0.5f, float size = 50f, float blur = 3f);

    /// <summary>Gets a shape's reflection visibility, transparency, size, and blur.</summary>
    ShapeOperationResult GetReflection(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets a shape's soft edge (feathered edge) radius in points. A radius of 0 removes the soft edge.</summary>
    ShapeOperationResult SetSoftEdge(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, float radius);

    /// <summary>Gets a shape's soft edge radius in points.</summary>
    ShapeOperationResult GetSoftEdge(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Applies a 3D bevel effect to a shape's top edge. <paramref name="bevelType"/> is an
    /// <c>MsoBevelType</c> enum member name (e.g. <c>"msoBevelCircle"</c>, <c>"msoBevelSoftRound"</c>,
    /// or <c>"msoBevelNone"</c> to remove the bevel).
    /// </summary>
    ShapeOperationResult SetBevel(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex, string bevelType, float depth = 6f, float inset = 6f);

    /// <summary>Gets a shape's bevel type name, depth, and inset.</summary>
    ShapeOperationResult GetBevel(ComInterop.Session.IPresentationBatch batch, int slideIndex, int shapeIndex);

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
