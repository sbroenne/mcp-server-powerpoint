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
                    ShapeIndex = FindShapeIndexById(shapes, newShape.Id),
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
            () => Interlocked.CompareExchange(ref callbackState, 2, 0) == 0,
            batch.PowerPointProcessIdentity);
        ClipboardLockCoordinator.Enqueue(request);

        // Polls in short slices rather than waiting out the whole budget in one call: dispatch
        // can run other commands against this same batch concurrently while this one sits queued
        // for the clipboard, so the session can be poisoned (or PowerPoint can exit) at any point
        // during the wait. Without polling, that would go unnoticed until this request finally
        // reached the front of the queue, wasting most of lockTimeout on a wait that was already
        // doomed. IsPowerPointProcessAlive() always reports false when no identity was captured
        // (nothing to confirm against), so it is only trusted as a "confirmed dead" signal when an
        // identity actually exists - otherwise a perfectly healthy session with no captured
        // identity would abandon every copy-to-slide call immediately.
        TimeSpan pollInterval = TimeSpan.FromMilliseconds(250);
        DateTime waitDeadlineUtc = DateTime.UtcNow + lockTimeout;
        while (!request.Acquired.Task.IsCompleted)
        {
            bool sessionUnusable = batch.HasTimedOutOperation
                || (batch.PowerPointProcessIdentity is not null && !batch.IsPowerPointProcessAlive());
            if (sessionUnusable)
            {
                request.MarkCallerGaveUp();
                throw new TimeoutException(
                    "The PowerPoint session became unusable while waiting for the shared clipboard " +
                    "(a previous operation timed out, or the PowerPoint process exited). The " +
                    "copy-to-slide operation was abandoned; open a new session and retry.");
            }

            TimeSpan remaining = waitDeadlineUtc - DateTime.UtcNow;
            if (remaining <= TimeSpan.Zero)
            {
                request.MarkCallerGaveUp();
                throw new TimeoutException(
                    $"Timed out after {lockTimeout.TotalSeconds:0.##} seconds waiting for the Windows " +
                    "clipboard, which another session or process is using to copy a shape. Retry the " +
                    "copy-to-slide operation.");
            }

            Task.WaitAny([request.Acquired.Task], remaining < pollInterval ? remaining : pollInterval);
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
                    // clipboard is no longer ours to touch. Still signal completion here (not
                    // just in the outer finally): this branch can run after the coordinator's own
                    // TryAbandon() already flipped callbackState away from this method's outer
                    // finally, in which case that finally's own CompareExchange no longer succeeds
                    // either, and nothing else would ever complete this task.
                    transferFinished.TrySetResult(true);
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
                    if (targetSlideIndex == slideIndex)
                    {
                        return new ShapeOperationResult
                        {
                            Success = false,
                            ErrorMessage = "targetSlideIndex must differ from slideIndex; use " +
                                "'duplicate' to copy a shape onto the same slide."
                        };
                    }

                    sourceSlide = slides[slideIndex];
                    sourceShapes = sourceSlide.Shapes;

                    var shapeValidation = ValidateShapeIndex(sourceShapes.Count, shapeIndex);
                    if (shapeValidation is not null) return shapeValidation;

                    sourceShape = sourceShapes[shapeIndex];
                    sourceShape.Copy();

                    targetSlide = slides[targetSlideIndex];
                    targetShapes = targetSlide.Shapes;
                    pasted = targetShapes.Paste();

                    // Best-effort clipboard-integrity check: this lock only serializes calls made
                    // through this server, not an unrelated application or a manual Ctrl+C/Ctrl+V
                    // on the same desktop session. If something else replaced the clipboard between
                    // Copy() and Paste(), PowerPoint pastes whatever is actually on it - often zero
                    // or more than one shape - rather than throwing. A count other than exactly one
                    // is a reliable signal of that, even though a same-count substitution cannot be
                    // detected this way.
                    if (pasted.Count != 1)
                    {
                        // Leaving whatever was actually pasted on the target slide would let a
                        // caller retrying after this failure accumulate unrelated content, so the
                        // partial mutation is undone before reporting it.
                        int pastedCount = pasted.Count;
                        pasted.Delete();
                        return new ShapeOperationResult
                        {
                            Success = false,
                            ErrorMessage = $"Clipboard integrity check failed: expected exactly 1 " +
                                $"shape after paste, got {pastedCount}. Another application or a " +
                                "manual copy/paste may have used the clipboard concurrently."
                        };
                    }

                    newShape = pasted[1];

                    return new ShapeOperationResult
                    {
                        Success = true,
                        ShapeIndex = FindShapeIndexById(targetShapes, newShape.Id),
                        ShapeCount = targetShapes.Count
                    };
                }
                finally
                {
                    // Reverse acquisition order (newShape/pasted/targetShapes/targetSlide were all
                    // acquired after sourceShape), matching CopyFormatting's release pattern.
                    if (newShape is not null) ComUtilities.Release(ref newShape);
                    if (pasted is not null) ComUtilities.Release(ref pasted);
                    if (targetShapes is not null) ComUtilities.Release(ref targetShapes);
                    if (targetSlide is not null) ComUtilities.Release(ref targetSlide);
                    if (sourceShape is not null) ComUtilities.Release(ref sourceShape);
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

            if (targetSlideIndex == slideIndex)
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = "targetSlideIndex must differ from slideIndex; use " +
                        "'duplicate' to copy a shape onto the same slide."
                };
            }

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
    /// Returns the 1-based index of the shape identified by <paramref name="shapeId"/> within
    /// <paramref name="shapes"/> - the same indexing every other shape command uses via
    /// <c>Shapes[index]</c>. <see cref="PowerPoint.Shape.ZOrderPosition"/> is deliberately not
    /// used for this: it tracks the shape's position on the z-order plane, which can diverge from
    /// its position in the <c>Shapes</c> collection after z-order changes, while <c>Shape.Id</c>
    /// is a stable per-presentation identifier safe to search on immediately after the shape is
    /// created.
    /// </summary>
    private static int FindShapeIndexById(PowerPoint.Shapes shapes, int shapeId)
    {
        for (int i = 1; i <= shapes.Count; i++)
        {
            PowerPoint.Shape? candidate = null;
            try
            {
                candidate = shapes[i];
                if (candidate.Id == shapeId)
                {
                    return i;
                }
            }
            finally
            {
                if (candidate is not null) ComUtilities.Release(ref candidate);
            }
        }

        throw new InvalidOperationException(
            $"Shape with Id {shapeId} was not found in the Shapes collection immediately after " +
            "it was created.");
    }

    /// <summary>
    /// A caller's turn at the clipboard lock. <paramref name="TryAbandon"/> returns true when the
    /// transfer had not started yet and can therefore never run, which is what makes releasing the
    /// lock safe after the caller has given up. <paramref name="PowerPointProcessIdentity"/>
    /// identifies (by PID and creation time, so PID reuse cannot transfer ownership to an
    /// unrelated process) the PowerPoint process this request's own transfer would run against.
    /// It is recorded only so that, if this request is the one that later turns out to be wedged,
    /// the coordinator can tell once that specific process is confirmed dead and safely clear the
    /// quarantine for every caller - it does not exempt this request's own process from a
    /// quarantine already in effect, because the clipboard being guarded is one desktop-wide
    /// resource, not a per-process one. It is <see langword="null"/> when the batch had not yet
    /// captured the identity of its owned process.
    /// </summary>
    internal sealed record ClipboardLockRequest(
        DateTime DeadlineUtc,
        TimeSpan CompletionTimeout,
        Func<bool> TryAbandon,
        PowerPointProcessIdentity? PowerPointProcessIdentity)
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
    internal static class ClipboardLockCoordinator
    {
        private static readonly System.Collections.Concurrent.BlockingCollection<ClipboardLockRequest> Requests = new();
        private static readonly Lock StartGate = new();
        private static bool s_started;

        // Non-null while a previous transfer never confirmed completion within its bounded waits
        // and may still be running. Read/written only via Volatile: Enqueue() runs on arbitrary
        // caller threads, while the field is cleared from the single coordinator thread once the
        // recorded process is confirmed dead, so both the wedge state and the identity it carries
        // must become visible together, not as separately-torn fields.
        private static volatile WedgeState? s_wedge;

        // The mutex handle retained (never released/disposed) for as long as a wedge is in
        // effect. Only ever touched by the single coordinator thread - the one thread .NET's
        // Mutex allows to release it - so no synchronization is needed here.
        private static Mutex? s_wedgeMutex;

        // TransferFinished is included so a wedged transfer that was merely slow - not truly stuck
        // - still resolves the quarantine as soon as its callback's finally block signals
        // completion, rather than only ever resolving once the recorded process exits.
        private sealed record WedgeState(PowerPointProcessIdentity? Identity, Task<bool> TransferFinished);

        // How often the coordinator thread re-checks a standing quarantine when no new request
        // has arrived to trigger the check itself (see Run()). Recovery must not depend on some
        // other caller happening to try again: this is a session-scoped named mutex, so a
        // different MCP/CLI process could be the one waiting on it while this process sits idle.
        private static readonly TimeSpan WedgeRecoveryPollInterval = TimeSpan.FromSeconds(15);

        // Test-only: set by ResetForTests() and consumed by Run() on the coordinator thread -
        // the only thread allowed to release s_wedgeMutex. See ResetForTests() for why this
        // indirection exists instead of releasing directly from the calling (test) thread.
        private static TaskCompletionSource<bool>? s_pendingTestReset;

        // Deliberately does NOT reject on IsQuarantineStillInEffect() here. Serve() - the only
        // place that calls TryClearResolvedQuarantine() - runs exclusively on the single
        // coordinator thread pulling from Requests, so a request that never reaches Requests can
        // never give that thread a chance to notice the wedge has resolved. Rejecting upfront
        // would make every quarantine permanent regardless of whether the recorded process later
        // exits: nothing would ever queue again to trigger the recheck. Every request is therefore
        // queued unconditionally, and Serve() itself performs the check-and-clear-if-resolved,
        // then re-checks before deciding whether to actually reject this particular request.
        // Run()'s own idle polling (see below) covers the case where no request ever arrives at
        // all after a wedge.
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

        // Runs for the lifetime of the process on a single dedicated thread. Normally blocks
        // indefinitely for the next request, same as a plain consuming enumerator. While a
        // quarantine is in effect, instead polls on WedgeRecoveryPollInterval so recovery does not
        // depend on some other request arriving to trigger TryClearResolvedQuarantine(): without
        // this, a wedge that resolves while this process is otherwise idle would hold the
        // session-scoped named mutex forever, blocking every other process waiting on it too.
        private static void Run()
        {
            while (true)
            {
                TaskCompletionSource<bool>? pendingReset = Interlocked.Exchange(ref s_pendingTestReset, null);
                if (pendingReset is not null)
                {
                    if (s_wedgeMutex is not null)
                    {
                        s_wedgeMutex.ReleaseMutex();
                        s_wedgeMutex.Dispose();
                        s_wedgeMutex = null;
                    }

                    s_wedge = null;
                    pendingReset.TrySetResult(true);
                }

                TimeSpan waitTimeout = IsQuarantineStillInEffect()
                    ? WedgeRecoveryPollInterval
                    : Timeout.InfiniteTimeSpan;

                if (Requests.TryTake(out var request, waitTimeout))
                {
                    Serve(request);
                }
                else if (IsQuarantineStillInEffect())
                {
                    TryClearResolvedQuarantine();
                }
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

            // A request can already be queued (enqueued while healthy) by the time an earlier
            // transfer discovers a wedge. Without this check, this coordinator thread would still
            // be the OS-recognized owner of the never-released named mutex below, so its own next
            // WaitOne() on that same name would re-acquire immediately (mutex ownership is
            // per-thread and recursive) and let this request run concurrently with the still-
            // possibly-active wedged transfer - exactly what the quarantine exists to prevent.
            // TryClearResolvedQuarantine runs first and, if the recorded process is now confirmed
            // dead, releases the retained mutex from this same thread and clears the quarantine
            // before this request is served - the only point at which doing so is both possible
            // (thread-affinity) and safe (the danger is actually gone, not just timed out on).
            if (!TryClearResolvedQuarantine() && IsQuarantineStillInEffect())
            {
                request.Acquired.TrySetException(QuarantineException());
                return;
            }

            Mutex? clipboardMutex = null;
            bool lockTaken = false;
            bool wedged = false;
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
                        // unlocking safe. If it still has not completed after this second bounded
                        // wait, treat it as wedged for now - it may simply be slow rather than
                        // truly stuck, so the quarantine recorded below still watches this same
                        // TransferFinished task and clears as soon as it completes, in addition to
                        // the process-exit fallback for a callback that never completes at all.
                        bool completed = request.TransferFinished.Task.Wait(request.CompletionTimeout);
                        if (!completed)
                        {
                            wedged = true;
                            s_wedge = new WedgeState(request.PowerPointProcessIdentity, request.TransferFinished.Task);
                        }
                    }
                }
            }
            finally
            {
                // ClipboardMutexName is a named system mutex, visible to every process in this
                // Windows session, not just this one - releasing it after a wedge would let a
                // different process's WaitOne() acquire it and run Copy()/Paste() while the
                // original STA callback may still be active, defeating the quarantine just as much
                // as this thread re-acquiring it would. So once wedged, this handle is deliberately
                // retained rather than released or disposed here: .NET's Mutex is thread-affine on
                // release, and Windows keeps the underlying kernel object owned by this thread for
                // as long as the thread itself lives, which for this dedicated background thread is
                // the rest of the process's lifetime unless TryClearResolvedQuarantine later
                // determines it is safe to let go. Every other process's own WaitOne() then blocks
                // or times out normally against a lock that is genuinely still held, instead of
                // acquiring one that only looks clean.
                if (wedged)
                {
                    s_wedgeMutex = clipboardMutex;
                }
                else
                {
                    if (lockTaken)
                    {
                        clipboardMutex!.ReleaseMutex();
                    }

                    clipboardMutex?.Dispose();
                }
            }
        }

        private static bool IsQuarantineStillInEffect() => s_wedge is not null;

        /// <summary>
        /// If a wedge is recorded and it can now be resolved, releases the retained mutex from
        /// this coordinator thread and clears the quarantine, returning <see langword="true"/>.
        /// Resolution happens either way:
        /// <list type="bullet">
        /// <item>The wedged request's own <see cref="ClipboardLockRequest.TransferFinished"/> task
        /// completes - the callback that claimed the clipboard was merely slow, not stuck, and its
        /// <c>finally</c> block signals completion once its real Copy()/Paste() call actually
        /// returns. This is the common, fast path and needs nothing further: the transfer that
        /// caused the wedge is verifiably done.</item>
        /// <item>A captured <see cref="PowerPointProcessIdentity"/> for which
        /// <see cref="OwnedProcessGuard.TryConfirmExited"/> confirms that exact process (matched by
        /// PID and creation time, so PID reuse cannot fool this) is gone - a fallback for a
        /// callback that truly never completes, since a dead process can never finish or still be
        /// running it.</item>
        /// </list>
        /// When no identity was captured for the wedged request and its transfer never completes,
        /// there is no way to ever confirm the danger is over, so the quarantine fails closed and
        /// never clears: a time-based fallback would let a later request run Copy()/Paste()
        /// concurrently with a COM call that, for all this coordinator can prove, might still be
        /// active, corrupting the shared clipboard or pasting into the wrong presentation - exactly
        /// what the quarantine exists to prevent. Only ever called from the single coordinator
        /// thread, which is the only thread .NET's Mutex allows to release the handle this method
        /// may dispose of.
        /// </summary>
        private static bool TryClearResolvedQuarantine()
        {
            WedgeState? wedge = s_wedge;
            if (wedge is null)
            {
                return false;
            }

            bool resolved = wedge.TransferFinished.IsCompleted
                || (wedge.Identity is PowerPointProcessIdentity identity && OwnedProcessGuard.TryConfirmExited(identity));
            if (!resolved)
            {
                return false;
            }

            if (s_wedgeMutex is not null)
            {
                s_wedgeMutex.ReleaseMutex();
                s_wedgeMutex.Dispose();
                s_wedgeMutex = null;
            }

            s_wedge = null;
            return true;
        }

        private static InvalidOperationException QuarantineException() => new(
            "The shared shape clipboard is quarantined: a previous copy-to-slide operation did " +
            "not confirm completion within its timeout. Handing out the lock now could race that " +
            "unknown, possibly still-active Copy()/Paste() pair. This clears automatically once " +
            "that operation actually completes (even if it was merely slow, not stuck) or, failing " +
            "that, once the affected PowerPoint process is confirmed to have exited. If that " +
            "process could not be identified when the wedge occurred, this quarantine cannot clear " +
            "automatically and the host process must be restarted.");

        // Test-only escape hatch: safely undoes a quarantine that this coordinator's own design
        // deliberately makes unrecoverable when no PowerPointProcessIdentity was captured (see
        // TryClearResolvedQuarantine's fail-closed branch). Real callers must never need this - a
        // genuine no-identity wedge really does require restarting the host process - but a test
        // that deliberately forces this state to prove it fails closed must not leave this shared
        // static coordinator permanently quarantined for every later test sharing this process.
        // The actual release/dispose happens on Run()'s coordinator thread (see the
        // s_pendingTestReset handling there), not here: Mutex.ReleaseMutex() requires the exact
        // thread that acquired ownership, and this method runs on the calling (test) thread.
        // Disposing the retained handle from here instead - without a matching ReleaseMutex() -
        // would leave the named mutex's OS-level ownership permanently held by the coordinator
        // thread with an unbalanced recursion count, silently blocking every later real holder.
        internal static void ResetForTests()
        {
            EnsureStarted();
            var pendingReset = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            Volatile.Write(ref s_pendingTestReset, pendingReset);

            // Wakes Run() immediately instead of waiting for its next poll: served as a no-op via
            // the CallerGaveUp early-return, so it never touches the mutex itself.
            var wakeRequest = new ClipboardLockRequest(DateTime.UtcNow, TimeSpan.Zero, () => true, null);
            wakeRequest.MarkCallerGaveUp();
            Requests.Add(wakeRequest);

            if (!pendingReset.Task.Wait(TimeSpan.FromSeconds(10)))
            {
                throw new TimeoutException("ClipboardLockCoordinator.ResetForTests timed out waiting for the coordinator thread.");
            }
        }
    }
}
