using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    private const string FormattingClipboardMutexName =
        "Sbroenne.PowerPointMcp.ShapeFormattingClipboard";

    // PowerPoint's Format Painter state is per-user global rather than per presentation, so the
    // PickUp/Apply pair has to be serialized across sessions AND processes; the MCP server and the
    // CLI daemon are separate processes, which is why this is not limited to the current session.
    private static NamedWaitHandleOptions FormattingClipboardMutexOptions => new()
    {
        CurrentUserOnly = true,
        CurrentSessionOnly = false
    };

    /// <inheritdoc/>
    public ShapeOperationResult CopyFormatting(
        IPresentationBatch batch,
        int slideIndex,
        int sourceShapeIndex,
        int targetShapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        // Deliberately taken on the calling thread rather than inside batch.Execute: the batch's
        // single STA thread must keep pumping COM messages, and its operation timeout only stops the
        // caller waiting - it cannot cancel a callback that is already blocked, which would strand
        // the thread and poison the session. Mutexes are thread-affine and Execute runs
        // synchronously, so this thread owns the lock for the whole transfer and releases it below.
        using var formattingClipboardMutex = new Mutex(
            FormattingClipboardMutexName,
            FormattingClipboardMutexOptions);
        bool lockTaken = false;
        try
        {
            // Half the caller's budget, so contending for the lock cannot double the worst-case
            // latency of the operation: the remainder stays available for the COM work itself.
            TimeSpan lockTimeout = batch.OperationTimeout / 2;
            try
            {
                lockTaken = formattingClipboardMutex.WaitOne(lockTimeout);
            }
            catch (AbandonedMutexException)
            {
                // The previous owner died without releasing; ownership transfers to this thread.
                lockTaken = true;
            }

            if (!lockTaken)
            {
                throw new TimeoutException(
                    $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for PowerPoint's " +
                    "formatting clipboard, which another session or process is using to copy shape " +
                    "formatting. Retry the copy-formatting operation.");
            }

            return batch.Execute((ctx, ct) =>
            {
                PowerPoint.Slides? slides = null;
                PowerPoint.Slide? slide = null;
                PowerPoint.Shapes? shapes = null;
                PowerPoint.Shape? sourceShape = null;
                PowerPoint.Shape? targetShape = null;
                try
                {
                    slides = ctx.Presentation.Slides;
                    var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                    if (slideValidation is not null) return slideValidation;

                    slide = slides[slideIndex];
                    shapes = slide.Shapes;

                    var sourceValidation = ValidateShapeIndex(shapes.Count, sourceShapeIndex);
                    if (sourceValidation is not null) return sourceValidation;

                    var targetValidation = ValidateShapeIndex(shapes.Count, targetShapeIndex);
                    if (targetValidation is not null) return targetValidation;

                    sourceShape = shapes[sourceShapeIndex];
                    targetShape = shapes[targetShapeIndex];
                    sourceShape.PickUp();
                    targetShape.Apply();

                    return new ShapeOperationResult
                    {
                        Success = true,
                        ShapeIndex = targetShapeIndex
                    };
                }
                finally
                {
                    if (targetShape is not null) ComUtilities.Release(ref targetShape);
                    if (sourceShape is not null) ComUtilities.Release(ref sourceShape);
                    if (shapes is not null) ComUtilities.Release(ref shapes);
                    if (slide is not null) ComUtilities.Release(ref slide);
                    if (slides is not null) ComUtilities.Release(ref slides);
                }
            });
        }
        finally
        {
            if (lockTaken)
            {
                formattingClipboardMutex.ReleaseMutex();
            }
        }
    }
}
