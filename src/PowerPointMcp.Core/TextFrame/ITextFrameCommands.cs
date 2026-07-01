using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>
/// Text frame commands: set/get text and basic font formatting (size, bold, color) for a
/// shape's text range. Operates within an already-open IPresentationBatch, targeting a
/// specific shape by its 1-based slide and shape index.
/// </summary>
public interface ITextFrameCommands
{
    /// <summary>Sets the text content of a shape's text frame.</summary>
    TextFrameOperationResult SetText(IPresentationBatch batch, int slideIndex, int shapeIndex, string text);

    /// <summary>Gets the text content of a shape's text frame.</summary>
    TextFrameOperationResult GetText(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font size (in points) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex, float fontSize);

    /// <summary>Sets whether a shape's entire text range is bold.</summary>
    TextFrameOperationResult SetBold(IPresentationBatch batch, int slideIndex, int shapeIndex, bool bold);

    /// <summary>Sets the font color (RGB) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue);
}
