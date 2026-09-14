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

        // Run the batch's own disposed / poisoned-session / PowerPoint-liveness checks before taking
        // the lock, so a broken session fails fast with its canonical error instead of first waiting
        // out the formatting-lock timeout.
        batch.Execute((ctx, ct) => { });

        // Half the caller's budget, so contending for the lock cannot double the worst-case
        // latency of the operation: the remainder stays available for the COM work itself.
        TimeSpan lockTimeout = batch.OperationTimeout / 2;

        // TaskCompletionSource rather than wait handles: the owner thread can outlive this call, so
        // anything disposable here would either leak OS handles or fault that thread.
        var lockResolved = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var transferFinished = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        // 0 = transfer not started, 1 = claimed by the callback, 2 = abandoned by the caller.
        int callbackState = 0;

        // The lock is owned by a dedicated thread: mutexes are thread-affine, the batch's STA thread
        // must stay free to pump COM messages, and batch.Execute can hand control back on its own
        // operation timeout while the queued callback is still running - releasing at that point
        // would let another session interleave between PickUp() and Apply().
        var lockOwner = new Thread(() =>
        {
            Mutex? formattingClipboardMutex = null;
            bool lockTaken = false;
            try
            {
                try
                {
                    formattingClipboardMutex = new Mutex(
                        FormattingClipboardMutexName,
                        FormattingClipboardMutexOptions);
                    lockTaken = formattingClipboardMutex.WaitOne(lockTimeout);
                    lockResolved.SetResult(lockTaken);
                }
                catch (AbandonedMutexException)
                {
                    // WaitOne throws this *after* transferring ownership, so the lock is held.
                    lockTaken = true;
                    lockResolved.SetResult(true);
                }
                catch (Exception ex)
                {
                    // Surface on the caller's thread instead of killing the process, and never
                    // leave the caller waiting for a signal that can no longer arrive.
                    lockResolved.SetException(ex);
                }

                if (lockTaken)
                {
                    transferFinished.Task.GetAwaiter().GetResult();
                    formattingClipboardMutex!.ReleaseMutex();
                }
            }
            finally
            {
                formattingClipboardMutex?.Dispose();
            }
        })
        {
            IsBackground = true,
            Name = "PowerPointFormattingClipboardLock"
        };

        lockOwner.Start();

        if (!lockResolved.Task.GetAwaiter().GetResult())
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
                if (Interlocked.CompareExchange(ref callbackState, 1, 0) != 0)
                {
                    // The caller already timed out and handed the lock back, so the shared
                    // formatting clipboard is no longer ours to touch.
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = "The copy-formatting operation was abandoned before it reached PowerPoint."
                    };
                }

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
                    transferFinished.TrySetResult(true);
                }
            });
        }
        finally
        {
            // Hand the lock back only when the callback never claimed it - if it did, its own
            // finally signals completion, so an Execute that timed out cannot unlock a transfer
            // that is still between PickUp() and Apply().
            if (Interlocked.CompareExchange(ref callbackState, 2, 0) == 0)
            {
                transferFinished.TrySetResult(true);
            }
        }
    }
}
