using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Media;

/// <summary>
/// Media commands: insert audio or video files and inspect their PowerPoint media metadata.
/// Operates within an already-open presentation session using 1-based slide and shape indexes.
/// </summary>
[ServiceCategory("media", "Media")]
[McpTool("media", Title = "Media Operations", Destructive = true, Category = "content",
    Description = "Insert embedded or linked audio and video, and inspect media metadata on a slide.")]
public interface IMediaCommands
{
    /// <summary>
    /// Adds an audio or video file to a slide. Use <paramref name="linkToFile"/> = false and
    /// <paramref name="saveWithDocument"/> = true for embedded media. Set
    /// <paramref name="linkToFile"/> = true for linked media; keep
    /// <paramref name="saveWithDocument"/> = true to also retain media data in the presentation,
    /// or false for a link-only shape. The false/false combination is invalid.
    /// </summary>
    MediaOperationResult AddMedia(
        IPresentationBatch batch,
        int slideIndex,
        string mediaPath,
        bool linkToFile,
        bool saveWithDocument,
        float left,
        float top,
        float width,
        float height);

    /// <summary>
    /// Gets the typed PowerPoint media kind and geometry for a media shape at a 1-based shape index.
    /// </summary>
    MediaOperationResult GetMediaInfo(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex);
}
