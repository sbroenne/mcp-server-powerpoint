using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Image;

/// <inheritdoc cref="IImageCommands"/>
public sealed class ImageCommands : IImageCommands
{
    private const int MsoFalse = 0;
    private const int MsoTrue = -1;

    /// <inheritdoc/>
    public ImageOperationResult AddPicture(IPresentationBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(imagePath);

        // Up-front validation (not a suppressed exception, Rule 1b) — gives a clear error
        // message instead of a generic COMException from AddPicture.
        string fullImagePath = Path.GetFullPath(imagePath);
        if (!File.Exists(fullImagePath))
        {
            return new ImageOperationResult
            {
                Success = false,
                ErrorMessage = $"Image file not found: {fullImagePath}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int slideCount = ctx.Presentation.Slides.Count;
            if (slideIndex < 1 || slideIndex > slideCount)
            {
                return new ImageOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
                };
            }

            // Shapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            // — LinkToFile/SaveWithDocument are Microsoft.Office.Core.MsoTriState (office.dll)
            // typed, so called late-bound via dynamic with the raw int constants, same pattern
            // as elsewhere in this project. LinkToFile=False, SaveWithDocument=True embeds the
            // image directly in the .pptx rather than linking to the external file.
            dynamic slide = ctx.Presentation.Slides[slideIndex];
            slide.Shapes.AddPicture(fullImagePath, MsoFalse, MsoTrue, left, top, width, height);
            int newIndex = (int)slide.Shapes.Count; // always appended

            return new ImageOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = (int)slide.Shapes.Count
            };
        });
    }
}
