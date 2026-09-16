using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Shape;

public sealed partial class ShapeCommands
{
    internal const string ClipboardMutexName = "Sbroenne.PowerPointMcp.ShapeClipboard";

    // Shape.Copy()/Shapes.Paste() use the Windows clipboard, a single resource shared by every
    // process on the user's desktop session (not just PowerPoint or this server), so concurrent
    // copies must be serialized the same way CopyFormatting's Format Painter buffer is serialized
    // above. Unlike the Format Painter buffer, the clipboard is session-scoped, not user-wide.
    internal static NamedWaitHandleOptions ClipboardMutexOptions => new()
    {
        CurrentUserOnly = true,
        CurrentSessionOnly = true
    };

    /// <inheritdoc/>
    public ShapeOperationResult Duplicate(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.ShapeRange? duplicated = null;
            PowerPoint.Shape? newShape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                if (slideValidation is not null) return slideValidation;

                slide = slides[slideIndex];
                shapes = slide.Shapes;

                var shapeValidation = ValidateShapeIndex(shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = shapes[shapeIndex];
                duplicated = shape.Duplicate();
                newShape = duplicated[1];

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newShape.ZOrderPosition,
                    ShapeCount = shapes.Count
                };
            }
            finally
            {
                if (newShape is not null) ComUtilities.Release(ref newShape);
                if (duplicated is not null) ComUtilities.Release(ref duplicated);
                if (shape is not null) ComUtilities.Release(ref shape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult CopyToSlide(IPresentationBatch batch, int slideIndex, int shapeIndex, int targetSlideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        // Validate before taking the clipboard lock, same rationale as CopyFormatting: a bad index
        // or a broken session fails fast instead of first waiting out the lock. The transfer below
        // revalidates, because the shapes/slides can change while another session holds the lock.
        var validation = batch.Execute((ctx, ct) =>
            ValidateCopyToSlideTargets(ctx, slideIndex, shapeIndex, targetSlideIndex));
        if (validation is not null)
        {
            return validation;
        }

        TimeSpan lockTimeout = batch.OperationTimeout / 2;

        // 0 = transfer not started, 1 = claimed by the callback, 2 = abandoned by the caller.
        int callbackState = 0;

        var request = new ClipboardLockRequest(
            DateTime.UtcNow + lockTimeout,
            batch.OperationTimeout,
            () => Interlocked.CompareExchange(ref callbackState, 2, 0) == 0);
        ClipboardLockCoordinator.Enqueue(request);

        if (Task.WaitAny([request.Acquired.Task], lockTimeout) < 0)
        {
            request.MarkCallerGaveUp();
            throw new TimeoutException(
                $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for the Windows " +
                "clipboard, which another session or process is using to copy a shape. Retry the " +
                "copy-to-slide operation.");
        }

        if (!request.Acquired.Task.GetAwaiter().GetResult())
        {
            throw new TimeoutException(
                $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for the Windows " +
                "clipboard, which another session or process is using to copy a shape. Retry the " +
                "copy-to-slide operation.");
        }

        var transferFinished = request.TransferFinished;

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                if (Interlocked.CompareExchange(ref callbackState, 1, 0) != 0)
                {
                    // The caller already timed out and handed the lock back, so the shared
                    // clipboard is no longer ours to touch.
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = "The copy-to-slide operation was abandoned before it reached PowerPoint."
                    };
                }

                PowerPoint.Slides? slides = null;
                PowerPoint.Slide? sourceSlide = null;
                PowerPoint.Slide? targetSlide = null;
                PowerPoint.Shapes? sourceShapes = null;
                PowerPoint.Shapes? targetShapes = null;
                PowerPoint.Shape? sourceShape = null;
                PowerPoint.ShapeRange? pasted = null;
                PowerPoint.Shape? newShape = null;
                try
                {
                    slides = ctx.Presentation.Slides;
                    var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
                    if (slideValidation is not null) return slideValidation;
                    var targetValidation = ValidateSlideIndex(slides.Count, targetSlideIndex);
                    if (targetValidation is not null) return targetValidation;

                    sourceSlide = slides[slideIndex];
                    sourceShapes = sourceSlide.Shapes;

                    var shapeValidation = ValidateShapeIndex(sourceShapes.Count, shapeIndex);
                    if (shapeValidation is not null) return shapeValidation;

                    sourceShape = sourceShapes[shapeIndex];
                    sourceShape.Copy();

                    targetSlide = slides[targetSlideIndex];
                    targetShapes = targetSlide.Shapes;
                    pasted = targetShapes.Paste();
                    newShape = pasted[1];

                    return new ShapeOperationResult
                    {
                        Success = true,
                        ShapeIndex = newShape.ZOrderPosition,
                        ShapeCount = targetShapes.Count
                    };
                }
                finally
                {
                    if (newShape is not null) ComUtilities.Release(ref newShape);
                    if (pasted is not null) ComUtilities.Release(ref pasted);
                    if (sourceShape is not null) ComUtilities.Release(ref sourceShape);
                    if (targetShapes is not null) ComUtilities.Release(ref targetShapes);
                    if (targetSlide is not null) ComUtilities.Release(ref targetSlide);
                    if (sourceShapes is not null) ComUtilities.Release(ref sourceShapes);
                    if (sourceSlide is not null) ComUtilities.Release(ref sourceSlide);
                    if (slides is not null) ComUtilities.Release(ref slides);
                    transferFinished.TrySetResult(true);
                }
            });
        }
        finally
        {
            // Hand the lock back only when the callback never claimed it - if it did, its own
            // finally signals completion, so an Execute that timed out cannot unlock a transfer
            // that is still between Copy() and Paste().
            if (Interlocked.CompareExchange(ref callbackState, 2, 0) == 0)
            {
                transferFinished.TrySetResult(true);
            }
        }
    }

    /// <summary>
    /// Checks that the source slide, shape, and target slide all exist, returning the failure
    /// result to hand back to the caller, or <see langword="null"/> when the request is valid.
    /// </summary>
    private static ShapeOperationResult? ValidateCopyToSlideTargets(
        PresentationContext ctx,
        int slideIndex,
        int shapeIndex,
        int targetSlideIndex)
    {
        PowerPoint.Slides? slides = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            slides = ctx.Presentation.Slides;
            var slideValidation = ValidateSlideIndex(slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            var targetValidation = ValidateSlideIndex(slides.Count, targetSlideIndex);
            if (targetValidation is not null) return targetValidation;

            slide = slides[slideIndex];
            shapes = slide.Shapes;

            return ValidateShapeIndex(shapes.Count, shapeIndex);
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (slide is not null) ComUtilities.Release(ref slide);
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    /// <summary>
    /// A caller's turn at the clipboard lock. <paramref name="TryAbandon"/> returns true when the
    /// transfer had not started yet and can therefore never run, which is what makes releasing the
    /// lock safe after the caller has given up.
    /// </summary>
    private sealed record ClipboardLockRequest(
        DateTime DeadlineUtc,
        TimeSpan CompletionTimeout,
        Func<bool> TryAbandon)
    {
        private int _callerGaveUp;

        public TaskCompletionSource<bool> Acquired { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public TaskCompletionSource<bool> TransferFinished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>Signalled when the caller stops waiting, so the coordinator never blocks on a
        /// completion that can no longer arrive.</summary>
        public TaskCompletionSource<bool> CallerGone { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>True once the caller stopped waiting, so the lock must not be handed to it.</summary>
        public bool CallerGaveUp => Volatile.Read(ref _callerGaveUp) == 1;

        // Deliberately internal, not public: check-core-interface-completeness.ps1 scans *Commands.cs
        // for public methods without distinguishing nested types, and would demand this on
        // IShapeCommands. The type itself is private, so this costs nothing.
        internal void MarkCallerGaveUp()
        {
            Interlocked.Exchange(ref _callerGaveUp, 1);
            CallerGone.TrySetResult(true);
        }
    }

    /// <summary>
    /// Owns the Windows clipboard lock on a single background thread for the whole process. The
    /// lock is exclusive, so one owner can serve every caller; a thread per request would
    /// accumulate whenever the clipboard or PowerPoint is busy.
    /// </summary>
    private static class ClipboardLockCoordinator
    {
        private static readonly System.Collections.Concurrent.BlockingCollection<ClipboardLockRequest> Requests = new();
        private static readonly Lock StartGate = new();
        private static bool s_started;

        internal static void Enqueue(ClipboardLockRequest request)
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
                    Name = "PowerPointShapeClipboardLock"
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

        private static void Serve(ClipboardLockRequest request)
        {
            // The caller stopped waiting while this request sat in the queue, so nobody will run a
            // transfer for it - do not take the lock on its behalf.
            if (request.CallerGaveUp)
            {
                request.Acquired.TrySetResult(false);
                return;
            }

            Mutex? clipboardMutex = null;
            bool lockTaken = false;
            try
            {
                try
                {
                    clipboardMutex = new Mutex(ClipboardMutexName, ClipboardMutexOptions);

                    // The deadline is the caller's, captured before queueing, so time spent waiting
                    // behind an earlier request still counts against it.
                    TimeSpan wait = request.DeadlineUtc - DateTime.UtcNow;
                    lockTaken = wait > TimeSpan.Zero && clipboardMutex.WaitOne(wait);
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
                    // Wake on the transfer finishing OR on the caller giving up: the caller can
                    // abandon in the instant between the check above and the handoff, and nothing
                    // would ever signal completion for it.
                    int signalled = Task.WaitAny(
                        [request.TransferFinished.Task, request.CallerGone.Task],
                        request.CompletionTimeout);

                    // The transfer cannot legitimately outlive the batch's own operation timeout.
                    if (signalled != 0 && !request.TryAbandon())
                    {
                        // The callback claimed the clipboard, so only its completion can make
                        // unlocking safe. Bounded anyway: a COM call wedged past its own timeout is
                        // already poisoning that session, and a user-wide lock held forever would
                        // break copy-to-slide in every session and process.
                        request.TransferFinished.Task.Wait(request.CompletionTimeout);
                    }
                }
            }
            finally
            {
                // Reaching here means either the transfer finished, TryAbandon closed the door on a
                // callback that never claimed, or the bounded wait above expired while a claimed
                // callback was still wedged inside a hung COM call. That last case is a real,
                // accepted gap, not an oversight: .NET's Mutex is thread-affine on release, and this
                // coordinator thread - not the STA thread running the wedged callback - is the only
                // one legally allowed to call ReleaseMutex() on the handle it acquired. There is no
                // safe way to hand release back to the callback's own thread, so releasing here
                // after the bounded wait (rather than blocking this coordinator, and therefore every
                // future caller, forever) is the same tradeoff CopyFormatting's Format Painter lock
                // already makes for the identical constraint.
                if (lockTaken)
                {
                    clipboardMutex!.ReleaseMutex();
                }

                clipboardMutex?.Dispose();
            }
        }
    }
}
