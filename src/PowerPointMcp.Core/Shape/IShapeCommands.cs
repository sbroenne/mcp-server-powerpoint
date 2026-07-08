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
}
