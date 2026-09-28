extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    private static readonly Dictionary<string, Office.MsoMergeCmd> MergeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoMergeUnion"] = Office.MsoMergeCmd.msoMergeUnion,
        ["msoMergeCombine"] = Office.MsoMergeCmd.msoMergeCombine,
        ["msoMergeIntersect"] = Office.MsoMergeCmd.msoMergeIntersect,
        ["msoMergeSubtract"] = Office.MsoMergeCmd.msoMergeSubtract,
        ["msoMergeFragment"] = Office.MsoMergeCmd.msoMergeFragment,
    };

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1822:Mark members as static",
        Justification = "IShapeCommands requires this operation to remain an instance command method.")]
    public ShapeOperationResult Merge(IPresentationBatch batch, int slideIndex, IReadOnlyList<int> shapeIndexes, string mergeType)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(shapeIndexes);
        ArgumentNullException.ThrowIfNull(mergeType);

        return batch.Execute((ctx, ct) =>
        {
            if (!MergeTypes.TryGetValue(mergeType, out var mergeCmd))
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{mergeType}' is not a recognized MsoMergeCmd name (must be " +
                        "'msoMergeUnion', 'msoMergeCombine', 'msoMergeIntersect', 'msoMergeSubtract', " +
                        "or 'msoMergeFragment')."
                };
            }

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
                int shapeCountBefore = shapes.Count;

                if (shapeIndexes.Count < 2)
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"At least 2 shape indexes are required to merge (got {shapeIndexes.Count})."
                    };
                }

                if (shapeIndexes.Distinct().Count() != shapeIndexes.Count)
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = "Shape indexes must be unique."
                    };
                }

                foreach (var index in shapeIndexes)
                {
                    var validation = ValidateShapeIndex(shapeCountBefore, index);
                    if (validation is not null) return validation;
                }

                object[] indexArray = shapeIndexes.Select(i => (object)i).ToArray();
                range = shapes.Range(indexArray);
                range.MergeShapes(mergeCmd);

                int shapeCountAfter = shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    MergeTypeName = mergeType,
                    MergedShapeCount = shapeCountAfter - shapeCountBefore + shapeIndexes.Count,
                    ShapeCount = shapeCountAfter
                };
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
