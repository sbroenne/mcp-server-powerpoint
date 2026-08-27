extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    /// <inheritdoc/>
    public ShapeOperationResult GetLinkInfo(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return ExecuteOnLinkedPicture(
            batch,
            slideIndex,
            shapeIndex,
            linkFormat => CreateLinkInfoResult(shapeIndex, linkFormat));
    }

    /// <inheritdoc/>
    public ShapeOperationResult UpdateLink(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return ExecuteOnLinkedPicture(
            batch,
            slideIndex,
            shapeIndex,
            linkFormat =>
            {
                string sourceFullName = linkFormat.SourceFullName;
                if (string.IsNullOrWhiteSpace(sourceFullName) || !File.Exists(sourceFullName))
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Linked source file not found: {sourceFullName}."
                    };
                }

                linkFormat.Update();
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    HasLink = true,
                    LinkSourceFullName = sourceFullName
                };
            });
    }

    /// <inheritdoc/>
    public ShapeOperationResult BreakLink(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return ExecuteOnLinkedPicture(
            batch,
            slideIndex,
            shapeIndex,
            linkFormat =>
            {
                linkFormat.BreakLink();
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    HasLink = false
                };
            });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetLinkAutoUpdate(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        bool autoUpdate)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return ExecuteOnLinkedPicture(
            batch,
            slideIndex,
            shapeIndex,
            linkFormat =>
            {
                linkFormat.AutoUpdate = autoUpdate
                    ? PowerPoint.PpUpdateOption.ppUpdateOptionAutomatic
                    : PowerPoint.PpUpdateOption.ppUpdateOptionManual;
                return CreateLinkInfoResult(shapeIndex, linkFormat, autoUpdate);
            });
    }

    private static ShapeOperationResult ExecuteOnLinkedPicture(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        Func<PowerPoint.LinkFormat, ShapeOperationResult> operation)
    {
        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.LinkFormat? linkFormat = null;
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
                if (shape.Type != Office.MsoShapeType.msoLinkedPicture)
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not linked. Link operations require a linked picture shape."
                    };
                }

                linkFormat = shape.LinkFormat;
                return operation(linkFormat);
            }
            finally
            {
                ComUtilities.Release(ref linkFormat);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref slide);
                ComUtilities.Release(ref slides);
            }
        });
    }

    private static ShapeOperationResult CreateLinkInfoResult(
        int shapeIndex,
        PowerPoint.LinkFormat linkFormat,
        bool? autoUpdate = null)
    {
        return new ShapeOperationResult
        {
            Success = true,
            ShapeIndex = shapeIndex,
            HasLink = true,
            LinkSourceFullName = linkFormat.SourceFullName,
            LinkAutoUpdate = autoUpdate
        };
    }
}
