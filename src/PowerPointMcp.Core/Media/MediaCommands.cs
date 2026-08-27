using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Media;

/// <inheritdoc cref="IMediaCommands"/>
public sealed class MediaCommands : IMediaCommands
{
    // AddMediaObject2's LinkToFile/SaveWithDocument parameters are Microsoft.Office.Core
    // MsoTriState values. office.dll is intentionally not referenced, so these named constants
    // cross only that late-bound COM boundary.
    private const int MsoFalse = 0; // MsoTriState.msoFalse
    private const int MsoTrue = -1; // MsoTriState.msoTrue

    // Shape.Type returns Microsoft.Office.Core.MsoShapeType, which is also unavailable without
    // office.dll. Keep media classification late-bound while MediaType stays typed as PpMediaType.
    private const int MsoMedia = 16; // MsoShapeType.msoMedia

    /// <inheritdoc/>
    public MediaOperationResult AddMedia(
        IPresentationBatch batch,
        int slideIndex,
        string mediaPath,
        bool linkToFile,
        bool saveWithDocument,
        float left,
        float top,
        float width,
        float height)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(mediaPath);

        if (string.IsNullOrWhiteSpace(mediaPath))
        {
            return Failure("Media path must not be empty.");
        }

        if (!linkToFile && !saveWithDocument)
        {
            return Failure(
                "saveWithDocument must be true when linkToFile is false; otherwise PowerPoint has " +
                "no media data to retain.");
        }

        if (width <= 0 || height <= 0)
        {
            return Failure("Media width and height must both be greater than zero.");
        }

        string fullMediaPath;
        try
        {
            fullMediaPath = Path.GetFullPath(mediaPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return Failure($"Invalid media path: {ex.Message}");
        }

        if (!File.Exists(fullMediaPath))
        {
            return Failure($"Media file not found: {fullMediaPath}.");
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? mediaShape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null)
                {
                    return slideValidation;
                }

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                // Reason: AddMediaObject2 requires Office.MsoTriState parameters from office.dll,
                // so only this method boundary is late-bound using named MsoTriState constants.
                mediaShape = (PowerPoint.Shape)((dynamic)shapes).AddMediaObject2(
                    fullMediaPath,
                    linkToFile ? MsoTrue : MsoFalse,
                    saveWithDocument ? MsoTrue : MsoFalse,
                    left,
                    top,
                    width,
                    height);

                PowerPoint.PpMediaType mediaType = mediaShape.MediaType;
                return new MediaOperationResult
                {
                    Success = true,
                    ShapeIndex = shapes.Count,
                    ShapeCount = shapes.Count,
                    MediaTypeName = mediaType.ToString(),
                    SourcePath = fullMediaPath,
                    LinkToFile = linkToFile,
                    SaveWithDocument = saveWithDocument,
                    Left = mediaShape.Left,
                    Top = mediaShape.Top,
                    Width = mediaShape.Width,
                    Height = mediaShape.Height
                };
            }
            finally
            {
                if (mediaShape is not null) ComUtilities.Release(ref mediaShape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public MediaOperationResult GetMediaInfo(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null)
                {
                    return slideValidation;
                }

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                var shapeValidation = ValidateShapeIndex(shapes.Count, shapeIndex);
                if (shapeValidation is not null)
                {
                    return shapeValidation;
                }

                shape = shapes[shapeIndex];
                // Reason: Shape.Type returns Office.MsoShapeType from office.dll; compare the
                // late-bound integer value with the named MsoShapeType.msoMedia constant.
                if ((int)((dynamic)shape).Type != MsoMedia)
                {
                    return Failure(
                        $"Shape {shapeIndex} on slide {slideIndex} is not a media shape.");
                }

                PowerPoint.PpMediaType mediaType = shape.MediaType;
                return new MediaOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    ShapeCount = shapes.Count,
                    MediaTypeName = mediaType.ToString(),
                    Left = shape.Left,
                    Top = shape.Top,
                    Width = shape.Width,
                    Height = shape.Height
                };
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    private static MediaOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return Failure(
                $"Slide index {slideIndex} is out of range. The presentation has {slideCount} " +
                $"slide(s) (valid range: 1-{slideCount}).");
        }

        return null;
    }

    private static MediaOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return Failure(
                $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} " +
                $"shape(s) (valid range: 1-{shapeCount}).");
        }

        return null;
    }

    private static MediaOperationResult Failure(string message)
        => new()
        {
            Success = false,
            ErrorMessage = message
        };
}
