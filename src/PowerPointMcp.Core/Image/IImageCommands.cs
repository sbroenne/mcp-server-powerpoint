using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Image commands: embed a picture file into a slide. Operates within an already-open
/// IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
public interface IImageCommands
{
    /// <summary>
    /// Adds a picture from a local file to the given slide, embedded (not linked) into
    /// the presentation.
    /// </summary>
    ImageOperationResult AddPicture(IPresentationBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height);
}
