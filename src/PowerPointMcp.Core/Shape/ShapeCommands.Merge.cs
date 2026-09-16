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

            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            int shapeCountBefore = slide.Shapes.Count;

            if (shapeIndexes.Count < 2)
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"At least 2 shape indexes are required to merge (got {shapeIndexes.Count})."
                };
            }

            foreach (var index in shapeIndexes)
            {
                var validation = ValidateShapeIndex(shapeCountBefore, index);
                if (validation is not null) return validation;
            }

            object[] indexArray = shapeIndexes.Select(i => (object)i).ToArray();
            PowerPoint.ShapeRange? range = null;
            try
            {
                range = slide.Shapes.Range(indexArray);
                range.MergeShapes(mergeCmd);
            }
            finally
            {
                if (range != null)
                {
                    ComUtilities.Release(ref range!);
                }
            }

            int shapeCountAfter = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                MergeTypeName = mergeType,
                MergedShapeCount = shapeCountAfter - shapeCountBefore + shapeIndexes.Count,
                ShapeCount = shapeCountAfter
            };
        });
    }
}
