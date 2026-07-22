using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>
/// Text frame commands: set/get text and basic font formatting (size, bold, italic, underline,
/// font name, color, alignment, bullets) for a shape's text range. Operates within an
/// already-open IPresentationBatch, targeting a specific shape by its 1-based slide and shape
/// index.
/// </summary>
[ServiceCategory("textframe", "TextFrame")]
[McpTool("textframe", Title = "Text Frame Operations", Destructive = true, Category = "content",
    Description = "Set or get text and font/paragraph formatting for a shape's text frame in an open presentation session.")]
public interface ITextFrameCommands
{
    /// <summary>Sets the text content of a shape's text frame.</summary>
    TextFrameOperationResult SetText(IPresentationBatch batch, int slideIndex, int shapeIndex, string text);

    /// <summary>Gets the text content of a shape's text frame.</summary>
    TextFrameOperationResult GetText(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font size (in points) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex, float fontSize);

    /// <summary>Gets the font size (in points) of a shape's text range.</summary>
    TextFrameOperationResult GetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is bold.</summary>
    TextFrameOperationResult SetBold(IPresentationBatch batch, int slideIndex, int shapeIndex, bool bold);

    /// <summary>Gets whether a shape's text range is bold.</summary>
    TextFrameOperationResult GetBold(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font color (RGB) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue);

    /// <summary>Gets the font color of a shape's text range as an RGB integer (0xBBGGRR, PowerPoint's native color order).</summary>
    TextFrameOperationResult GetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is italic.</summary>
    TextFrameOperationResult SetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex, bool italic);

    /// <summary>Gets whether a shape's text range is italic.</summary>
    TextFrameOperationResult GetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is underlined.</summary>
    TextFrameOperationResult SetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex, bool underline);

    /// <summary>Gets whether a shape's text range is underlined.</summary>
    TextFrameOperationResult GetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font name (typeface) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex, string fontName);

    /// <summary>Gets the font name (typeface) of a shape's text range.</summary>
    TextFrameOperationResult GetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the paragraph alignment of a shape's entire text range, identified by its
    /// <c>PpParagraphAlignment</c> enum member name (<c>"ppAlignLeft"</c>, <c>"ppAlignCenter"</c>,
    /// <c>"ppAlignRight"</c>, <c>"ppAlignJustify"</c>, or <c>"ppAlignDistribute"</c>).
    /// </summary>
    TextFrameOperationResult SetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex, string alignment);

    /// <summary>Gets the paragraph alignment of a shape's text range as a <c>PpParagraphAlignment</c> enum member name.</summary>
    TextFrameOperationResult GetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Turns bullets on or off for a shape's entire text range. When <paramref name="enabled"/> is
    /// <c>true</c>, an optional single-character <paramref name="character"/> sets the bullet
    /// glyph (defaults to PowerPoint's theme bullet character if omitted).
    /// </summary>
    TextFrameOperationResult SetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex, bool enabled, string? character = null);

    /// <summary>Gets whether a shape's text range has bullets enabled, and the bullet character if so.</summary>
    TextFrameOperationResult GetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets how a shape's text frame automatically resizes text/shape, identified by its
    /// <c>PpAutoSize</c> enum member name (<c>"ppAutoSizeNone"</c>, <c>"ppAutoSizeShapeToFitText"</c>,
    /// or <c>"ppAutoSizeTextToFitShape"</c>).
    /// </summary>
    TextFrameOperationResult SetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex, string autoSize);

    /// <summary>Gets a shape's text frame auto-size mode as a <c>PpAutoSize</c> enum member name.</summary>
    TextFrameOperationResult GetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex);
}
