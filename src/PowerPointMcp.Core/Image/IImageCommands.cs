using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <summary>
/// Image commands: embed a picture file into a slide. Operates within an already-open
/// IPresentationBatch, targeting a specific slide by its 1-based index.
/// </summary>
[ServiceCategory("image", "Image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "content",
    Description = "Embed a picture file into a slide in an open presentation session.")]
public interface IImageCommands
{
    /// <summary>
    /// Adds a picture from a local file to the given slide, embedded (not linked) into
    /// the presentation.
    /// </summary>
    ImageOperationResult AddPicture(IPresentationBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height);
}
