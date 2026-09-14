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

        // Half the caller's budget, so contending for the lock cannot double the worst-case
        // latency of the operation: the remainder stays available for the COM work itself.
        TimeSpan lockTimeout = batch.OperationTimeout / 2;

        // Not disposed deliberately: the owner thread below can outlive this call when a timed-out
        // Execute returns while the queued callback is still running, and disposing these from here
        // would fault that thread.
        var lockResolved = new ManualResetEventSlim(false);
        var transferFinished = new ManualResetEventSlim(false);
        bool lockTaken = false;
        int callbackState = 0;

        // The lock is owned by a dedicated thread: mutexes are thread-affine, the batch's STA thread
        // must stay free to pump COM messages, and batch.Execute can hand control back on its own
        // operation timeout while the queued callback is still running - releasing at that point
        // would let another session interleave between PickUp() and Apply().
        var lockOwner = new Thread(() =>
        {
            using var formattingClipboardMutex = new Mutex(
                FormattingClipboardMutexName,
                FormattingClipboardMutexOptions);
            try
            {
                lockTaken = formattingClipboardMutex.WaitOne(lockTimeout);
            }
            catch (AbandonedMutexException)
            {
                // The previous owner died without releasing; ownership transfers to this thread.
                lockTaken = true;
            }
            finally
            {
                lockResolved.Set();
            }

            if (!lockTaken)
            {
                return;
            }

            transferFinished.Wait();
            formattingClipboardMutex.ReleaseMutex();
        })
        {
            IsBackground = true,
            Name = "PowerPointFormattingClipboardLock"
        };

        lockOwner.Start();
        lockResolved.Wait();

        if (!lockTaken)
        {
            throw new TimeoutException(
                $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for PowerPoint's " +
                "formatting clipboard, which another session or process is using to copy shape " +
                "formatting. Retry the copy-formatting operation.");
        }

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                Interlocked.CompareExchange(ref callbackState, 1, 0);
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
                    transferFinished.Set();
                }
            });
        }
        finally
        {
            // Hand the lock back only when the callback never started - if it did start, its own
            // finally signals completion, so an Execute that timed out cannot unlock a transfer
            // that is still between PickUp() and Apply().
            if (Interlocked.CompareExchange(ref callbackState, 2, 0) == 0)
            {
                transferFinished.Set();
            }
        }
    }
}
