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

        // One budget spans validation, the lock wait and the COM work, so a contended call cannot
        // exceed the session's configured operation timeout.
        var elapsed = System.Diagnostics.Stopwatch.StartNew();
        TimeSpan totalBudget = batch.OperationTimeout;

        TimeSpan Remaining() => totalBudget - elapsed.Elapsed;

        TimeoutException BudgetExhausted(string phase) => new(
            $"Copy-formatting exceeded its {totalBudget.TotalSeconds:0.##} second budget during {phase}. " +
            "Retry the copy-formatting operation.");

        TResult RunWithinBudget<TResult>(
            string phase,
            Func<PresentationContext, CancellationToken, TResult> operation)
        {
            TimeSpan remaining = Remaining();
            if (remaining <= TimeSpan.Zero)
            {
                throw BudgetExhausted(phase);
            }

            using var budgetCts = new CancellationTokenSource(remaining);
            try
            {
                return batch.Execute(operation, budgetCts.Token);
            }
            catch (OperationCanceledException) when (budgetCts.IsCancellationRequested)
            {
                throw BudgetExhausted(phase);
            }
        }

        // Validate before taking the lock - this also runs the batch's disposed, poisoned-session and
        // PowerPoint-liveness checks - so a bad index or a broken session fails fast instead of first
        // waiting out the lock. The transfer below revalidates, because the shapes can change while
        // another session holds the lock.
        var validation = RunWithinBudget(
            "validation",
            (ctx, ct) => ValidateCopyFormattingTargets(ctx, slideIndex, sourceShapeIndex, targetShapeIndex));
        if (validation is not null)
        {
            return validation;
        }

        // At most half the budget goes on the lock, so a contended wait still leaves time for the
        // COM work, and never more than the budget actually has left.
        TimeSpan lockTimeout = TimeSpan.FromTicks(Math.Min(Remaining().Ticks, (totalBudget / 2).Ticks));
        if (lockTimeout <= TimeSpan.Zero)
        {
            throw BudgetExhausted("validation");
        }

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
                    // Bounded on purpose: if PowerPoint wedges mid-transfer the callback never signals,
                    // and holding a user-wide lock forever would break copy-formatting in every session
                    // and process. Only the budget the caller still has can justify holding it, so the
                    // lock is never kept past the point where the transfer itself has timed out.
                    TimeSpan releaseWait = Remaining();
                    transferFinished.Task.Wait(releaseWait < TimeSpan.Zero ? TimeSpan.Zero : releaseWait);
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
            return RunWithinBudget("the formatting transfer", (ctx, ct) =>
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

    /// <summary>
    /// Checks that the slide and both shape indexes exist, returning the failure result to hand back
    /// to the caller, or <see langword="null"/> when the request is valid.
    /// </summary>
    private static ShapeOperationResult? ValidateCopyFormattingTargets(
        PresentationContext ctx,
        int slideIndex,
        int sourceShapeIndex,
        int targetShapeIndex)
    {
        PowerPoint.Slides? slides = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            slides = ctx.Presentation.Slides;
            var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            slide = slides[slideIndex];
            shapes = slide.Shapes;

            var sourceValidation = ValidateShapeIndex(shapes.Count, sourceShapeIndex);
            if (sourceValidation is not null) return sourceValidation;

            return ValidateShapeIndex(shapes.Count, targetShapeIndex);
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (slide is not null) ComUtilities.Release(ref slide);
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }
}
