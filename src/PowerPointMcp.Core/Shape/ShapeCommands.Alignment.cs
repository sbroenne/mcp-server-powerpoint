extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    /// <inheritdoc/>
    public ShapeOperationResult Align(IPresentationBatch batch, int slideIndex,
        IReadOnlyList<int> shapeIndexes, string alignCmd, bool relativeToSlide = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (!Enum.TryParse<Office.MsoAlignCmd>(alignCmd, true, out var command) ||
            !Enum.IsDefined(command) || !string.Equals(alignCmd, command.ToString(), StringComparison.OrdinalIgnoreCase))
        {
            return new ShapeOperationResult { ErrorMessage = "alignCmd must be msoAlignLefts, msoAlignCenters, msoAlignRights, msoAlignTops, msoAlignMiddles, or msoAlignBottoms." };
        }

        return ArrangeShapes(batch, slideIndex, shapeIndexes, relativeToSlide ? 1 : 2,
            range => range.Align(command, relativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse));
    }

    /// <inheritdoc/>
    public ShapeOperationResult Distribute(IPresentationBatch batch, int slideIndex,
        IReadOnlyList<int> shapeIndexes, string distributeCmd, bool relativeToSlide = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (!Enum.TryParse<Office.MsoDistributeCmd>(distributeCmd, true, out var command) ||
            !Enum.IsDefined(command) || !string.Equals(distributeCmd, command.ToString(), StringComparison.OrdinalIgnoreCase))
        {
            return new ShapeOperationResult { ErrorMessage = "distributeCmd must be msoDistributeHorizontally or msoDistributeVertically." };
        }

        return ArrangeShapes(batch, slideIndex, shapeIndexes, 3,
            range => range.Distribute(command, relativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse));
    }

    private static ShapeOperationResult ArrangeShapes(IPresentationBatch batch, int slideIndex,
        IReadOnlyList<int>? shapeIndexes, int minimumCount, Action<PowerPoint.ShapeRange> arrange)
    {
        if (shapeIndexes is null || shapeIndexes.Count < minimumCount)
        {
            return new ShapeOperationResult { ErrorMessage = $"At least {minimumCount} distinct shape indexes are required." };
        }

        int[] indexes = shapeIndexes.ToArray();
        if (indexes.Distinct().Count() != indexes.Length)
        {
            return new ShapeOperationResult { ErrorMessage = "shapeIndexes must not contain duplicate indexes." };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.ShapeRange? range = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null) return slideValidation;

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                int shapeCount = shapes.Count;
                foreach (int index in indexes)
                {
                    var shapeValidation = ValidateShapeIndex(shapeCount, index);
                    if (shapeValidation is not null) return shapeValidation;
                }

                ct.ThrowIfCancellationRequested();
                range = shapes.Range(indexes.Select(index => (object)index).ToArray());
                arrange(range);
                return new ShapeOperationResult { Success = true, ShapeCount = shapeCount };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }
}