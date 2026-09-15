using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    internal const string FormattingClipboardMutexName =
        "Sbroenne.PowerPointMcp.ShapeFormattingClipboard";

    // PowerPoint's Format Painter state is per-user global rather than per presentation, so the
    // PickUp/Apply pair has to be serialized across sessions AND processes; the MCP server and the
    // CLI daemon are separate processes, which is why this is not limited to the current session.
    internal static NamedWaitHandleOptions FormattingClipboardMutexOptions => new()
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

        // Two-phase timeout policy, deliberately not one end-to-end budget:
        //   * every COM callback runs on the batch's own operation timeout, so an unresponsive
        //     PowerPoint hits the internal timeout and poisons the session like any other command;
        //   * only the lock wait, which touches no COM, uses the bounded non-poisoning wait below.
        // A contended call can therefore take the lock wait plus one operation timeout; detecting a
        // wedged COM call is worth more than a strict total-time guarantee.

        // Validate before taking the lock - this also runs the batch's disposed, poisoned-session and
        // PowerPoint-liveness checks - so a bad index or a broken session fails fast instead of first
        // waiting out the lock. The transfer below revalidates, because the shapes can change while
        // another session holds the lock.
        var validation = batch.Execute((ctx, ct) =>
            ValidateCopyFormattingTargets(ctx, slideIndex, sourceShapeIndex, targetShapeIndex));
        if (validation is not null)
        {
            return validation;
        }

        // Half the session's operation timeout is the most a caller should spend queueing behind
        // another session's transfer before being told to retry.
        TimeSpan lockTimeout = batch.OperationTimeout / 2;

        // 0 = transfer not started, 1 = claimed by the callback, 2 = abandoned by the caller.
        int callbackState = 0;

        // One shared coordinator owns the lock for every request in this process: mutexes are
        // thread-affine and the owner has to outlive an Execute that returned on its own timeout,
        // but a thread per request would pile up under a burst of contended calls.
        var request = new FormattingLockRequest(
            DateTime.UtcNow + lockTimeout,
            batch.OperationTimeout,
            () => Interlocked.CompareExchange(ref callbackState, 2, 0) == 0);
        FormattingLockCoordinator.Enqueue(request);

        // Bound the wait here as well as in the coordinator: an earlier request's transfer must not
        // be able to push this caller past its own deadline while it sits in the queue.
        if (!request.Acquired.Task.Wait(lockTimeout))
        {
            request.MarkCallerGaveUp();
            throw new TimeoutException(
                $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for PowerPoint's " +
                "formatting clipboard, which another session or process is using to copy shape " +
                "formatting. Retry the copy-formatting operation.");
        }

        if (!request.Acquired.Task.GetAwaiter().GetResult())
        {
            throw new TimeoutException(
                $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for PowerPoint's " +
                "formatting clipboard, which another session or process is using to copy shape " +
                "formatting. Retry the copy-formatting operation.");
        }

        var transferFinished = request.TransferFinished;

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

    /// <summary>
    /// A caller's turn at the formatting clipboard lock. <paramref name="TryAbandon"/> returns true
    /// when the transfer had not started yet and can therefore never run, which is what makes
    /// releasing the lock safe after the caller has given up.
    /// </summary>
    private sealed record FormattingLockRequest(
        DateTime DeadlineUtc,
        TimeSpan CompletionTimeout,
        Func<bool> TryAbandon)
    {
        private int _callerGaveUp;

        public TaskCompletionSource<bool> Acquired { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public TaskCompletionSource<bool> TransferFinished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>True once the caller stopped waiting, so the lock must not be handed to it.</summary>
        public bool CallerGaveUp => Volatile.Read(ref _callerGaveUp) == 1;

        public void MarkCallerGaveUp() => Interlocked.Exchange(ref _callerGaveUp, 1);
    }

    /// <summary>
    /// Owns PowerPoint's user-wide formatting clipboard lock on a single background thread for the
    /// whole process. The lock is exclusive, so one owner can serve every caller; a thread per
    /// request would accumulate whenever the clipboard or PowerPoint is busy.
    /// </summary>
    private static class FormattingLockCoordinator
    {
        private static readonly System.Collections.Concurrent.BlockingCollection<FormattingLockRequest> Requests = new();
        private static readonly Lock StartGate = new();
        private static bool s_started;

        internal static void Enqueue(FormattingLockRequest request)
        {
            EnsureStarted();
            Requests.Add(request);
        }

        private static void EnsureStarted()
        {
            if (Volatile.Read(ref s_started)) return;

            lock (StartGate)
            {
                if (s_started) return;

                new Thread(Run)
                {
                    IsBackground = true,
                    Name = "PowerPointFormattingClipboardLock"
                }.Start();
                Volatile.Write(ref s_started, true);
            }
        }

        private static void Run()
        {
            foreach (var request in Requests.GetConsumingEnumerable())
            {
                Serve(request);
            }
        }

        private static void Serve(FormattingLockRequest request)
        {
            // The caller stopped waiting while this request sat in the queue, so nobody will run a
            // transfer for it - do not take the lock on its behalf.
            if (request.CallerGaveUp)
            {
                request.Acquired.TrySetResult(false);
                return;
            }

            Mutex? formattingClipboardMutex = null;
            bool lockTaken = false;
            try
            {
                try
                {
                    formattingClipboardMutex = new Mutex(
                        FormattingClipboardMutexName,
                        FormattingClipboardMutexOptions);

                    // The deadline is the caller's, captured before queueing, so time spent waiting
                    // behind an earlier request still counts against it.
                    TimeSpan wait = request.DeadlineUtc - DateTime.UtcNow;
                    lockTaken = wait > TimeSpan.Zero && formattingClipboardMutex.WaitOne(wait);
                    request.Acquired.TrySetResult(lockTaken);
                }
                catch (AbandonedMutexException)
                {
                    // WaitOne throws this *after* transferring ownership, so the lock is held.
                    lockTaken = true;
                    request.Acquired.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    // Surface on the caller's thread instead of killing the process, and never leave
                    // the caller waiting for a signal that can no longer arrive.
                    request.Acquired.TrySetException(ex);
                }

                if (lockTaken)
                {
                    // The caller gave up between the deadline check and the handoff, so there is no
                    // transfer to protect.
                    if (request.CallerGaveUp)
                    {
                        return;
                    }

                    // The transfer cannot legitimately outlive the batch's own operation timeout.
                    if (!request.TransferFinished.Task.Wait(request.CompletionTimeout)
                        && !request.TryAbandon())
                    {
                        // The callback claimed the clipboard and has now outlived even its own
                        // timeout, so the session is already being poisoned by that timeout. Give the
                        // handoff one more bounded chance, then unlock regardless: a user-wide lock
                        // held forever would break copy-formatting in every session and process.
                        request.TransferFinished.Task.Wait(request.CompletionTimeout);
                    }
                }
            }
            finally
            {
                // Reaching here means either the transfer finished, or it can no longer start:
                // TryAbandon closes the door on a callback that never claimed.
                if (lockTaken)
                {
                    formattingClipboardMutex!.ReleaseMutex();
                }

                formattingClipboardMutex?.Dispose();
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
