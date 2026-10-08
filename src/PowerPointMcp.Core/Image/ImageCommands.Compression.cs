extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Image;

public sealed partial class ImageCommands
{
    /// <inheritdoc/>
    public ImageOperationResult CompressPictures(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? shapeIndex = null,
        string resolution = "print",
        bool deleteCroppedAreas = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(resolution);

        if (shapeIndex.HasValue && !slideIndex.HasValue)
        {
            return Failure("slideIndex is required when shapeIndex is provided.");
        }

        string normalizedResolution = resolution.Trim().ToLowerInvariant();
        (string? Name, int? Ppi) preset = normalizedResolution switch
        {
            "high-fidelity" => ("high-fidelity", null),
            "hd" => ("hd", 330),
            "print" => ("print", 220),
            "web" => ("web", 150),
            "email" => ("email", 96),
            _ => default
        };
        if (preset.Name is null)
        {
            return Failure(
                $"'{resolution}' is not a supported resolution. Use 'high-fidelity', 'hd', 'print', 'web', or 'email'.");
        }

        if (!OperatingSystem.IsWindows())
        {
            return Failure("Picture compression requires Windows.");
        }

        string extension = Path.GetExtension(batch.PresentationPath);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase))
        {
            return Failure("Picture compression requires a saved .pptx or .pptm presentation.");
        }

        var selection = batch.Execute((ctx, ct) =>
        {
            var selectedPictures = new Dictionary<int, HashSet<int>>();
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex.HasValue)
                {
                    ImageOperationResult? slideValidation = ValidateSlideIndex(slides.Count, slideIndex.Value);
                    if (slideValidation is not null)
                    {
                        return (selectedPictures, slideValidation);
                    }
                }

                int firstSlide = slideIndex ?? 1;
                int lastSlide = slideIndex ?? slides.Count;
                for (int currentSlideIndex = firstSlide; currentSlideIndex <= lastSlide; currentSlideIndex++)
                {
                    slide = slides[currentSlideIndex];
                    shapes = slide.Shapes;

                    if (shapeIndex.HasValue)
                    {
                        ImageOperationResult? shapeValidation = ValidateShapeIndex(shapes.Count, shapeIndex.Value);
                        if (shapeValidation is not null)
                        {
                            return (selectedPictures, shapeValidation);
                        }

                        shape = shapes[shapeIndex.Value];
                        ImageOperationResult? pictureValidation =
                            ValidatePictureShape(shape, currentSlideIndex, shapeIndex.Value);
                        if (pictureValidation is not null)
                        {
                            return (selectedPictures, pictureValidation);
                        }

                        selectedPictures[currentSlideIndex] = [shape.Id];
                        ComUtilities.Release(ref shape);
                    }
                    else
                    {
                        var pictureIds = new HashSet<int>();
                        for (int currentShapeIndex = 1; currentShapeIndex <= shapes.Count; currentShapeIndex++)
                        {
                            shape = shapes[currentShapeIndex];
                            AddPictureShapeIds(shape, pictureIds);
                            ComUtilities.Release(ref shape);
                        }

                        if (pictureIds.Count > 0)
                        {
                            selectedPictures[currentSlideIndex] = pictureIds;
                        }
                    }

                    ComUtilities.Release(ref shapes);
                    ComUtilities.Release(ref slide);
                }

                return (selectedPictures, (ImageOperationResult?)null);
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
        });

        if (selection.Item2 is not null)
        {
            return selection.Item2;
        }
        if (selection.Item1.Count == 0)
        {
            return new ImageOperationResult
            {
                Success = true,
                Resolution = preset.Name,
                CompressedPictureCount = 0,
                SkippedPictureCount = 0,
                OriginalImageBytes = 0,
                CompressedImageBytes = 0,
                SkippedPictures = []
            };
        }

        PresentationCompressionResult compressed = batch.TransformPresentationCopy(
            (path, cancellationToken) => PresentationImageCompressor.Compress(
                path,
                selection.Item1,
                preset.Name,
                preset.Ppi,
                deleteCroppedAreas,
                cancellationToken));

        return new ImageOperationResult
        {
            Success = true,
            Resolution = compressed.Resolution,
            CompressedPictureCount = compressed.CompressedPictureCount,
            SkippedPictureCount = compressed.SkippedPictures.Count,
            OriginalImageBytes = compressed.OriginalImageBytes,
            CompressedImageBytes = compressed.CompressedImageBytes,
            SkippedPictures = compressed.SkippedPictures
        };
    }

    private static void AddPictureShapeIds(PowerPoint.Shape shape, HashSet<int> pictureIds)
    {
        if (shape.Type is Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoLinkedPicture)
        {
            pictureIds.Add(shape.Id);
            return;
        }

        if (shape.Type != Office.MsoShapeType.msoGroup)
        {
            return;
        }

        PowerPoint.GroupShapes? groupItems = null;
        PowerPoint.Shape? groupItem = null;
        try
        {
            groupItems = shape.GroupItems;
            for (int index = 1; index <= groupItems.Count; index++)
            {
                groupItem = groupItems[index];
                AddPictureShapeIds(groupItem, pictureIds);
                ComUtilities.Release(ref groupItem);
            }
        }
        finally
        {
            ComUtilities.Release(ref groupItem);
            ComUtilities.Release(ref groupItems);
        }
    }
}
