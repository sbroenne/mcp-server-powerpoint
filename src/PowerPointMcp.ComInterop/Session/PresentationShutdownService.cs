using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Polly;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Centralized service for PowerPoint presentation close and application quit operations.
/// Implements resilient shutdown with retry/backoff for COM busy conditions, ported from
/// mcp-server-excel's <c>ExcelShutdownService</c>.
/// </summary>
/// <remarks>
/// <para><b>Shutdown order</b> (called from <see cref="PresentationBatch"/>'s STA thread once its
/// message pump loop exits):</para>
/// <list type="number">
/// <item>Close() the presentation, retrying transient COM-busy errors
/// (RPC_E_SERVERCALL_RETRYLATER, RPC_E_CALL_REJECTED) with linear backoff.</item>
/// <item>Release the presentation COM reference.</item>
/// <item>Quit() the application, retrying transient COM-busy errors with exponential backoff +
/// jitter (<see cref="ResiliencePipelines.CreatePowerPointQuitPipeline"/>).</item>
/// <item>Release the application COM reference.</item>
/// <item>If a process ID was captured, poll <c>Process.HasExited</c> with exponential backoff
/// for up to <see cref="ComInteropConstants.PowerPointProcessExitGracePeriod"/> before
/// force-terminating as a last resort.</item>
/// </list>
/// <para>
/// <b>Why the process-exit poll matters (Ripley's finding, recorded 2026-07-01):</b> a live
/// MCP round-trip integration test observed POWERPNT.exe legitimately taking ~90-100 seconds to
/// leave the OS process list after <c>Quit()</c> returns — Office's own COM/telemetry cleanup,
/// not a hang. <see cref="ComInteropConstants.PowerPointProcessExitGracePeriod"/> is sized
/// comfortably above that window so the happy path NEVER force-kills. Force-termination is
/// reserved strictly for the case where PowerPoint is still alive after the full grace period
/// elapses (a genuine hang, e.g. a stuck modal dialog or dead RPC connection).
/// </para>
/// <para>
/// This runs on the batch's dedicated STA thread as part of its own cleanup.
/// <c>PresentationBatch.Dispose()</c> deliberately joins that thread with
/// <see cref="ComInteropConstants.StaThreadJoinTimeout"/>, which is sized to cover this
/// service's full worst-case duration (Close()/Quit() retries + the process-exit grace period +
/// force-kill) — NOT just the Quit() call. A shorter join timeout would let Dispose() (and the
/// host process) return/exit before the grace-period poll or force-kill escalation ever runs,
/// silently defeating the safety net and leaking POWERPNT.exe (this was tried and observed to
/// leak indefinitely — see history.md). On the happy path (PowerPoint quits within seconds),
/// <c>Thread.Join()</c> returns as soon as the STA thread actually finishes, so this does not
/// slow down normal shutdown.
/// </para>
/// </remarks>
internal static class PresentationShutdownService
{
    private static readonly ResiliencePipeline QuitPipeline = ResiliencePipelines.CreatePowerPointQuitPipeline();

    private const int CloseMaxAttempts = 3;
    private const int CloseRetryDelayMs = 200;

    private const int InitialPollDelayMs = 250;
    private const int MaxPollDelayMs = 8000;

    /// <summary>
    /// Closes the presentation and quits the PowerPoint application with resilient retry logic,
    /// then polls (with exponential backoff) for the process to actually exit before ever
    /// force-terminating it.
    /// </summary>
    /// <param name="presentation">Presentation COM object (can be null).</param>
    /// <param name="app">Application COM object (can be null).</param>
    /// <param name="processId">
    /// POWERPNT.exe process ID captured at startup, if any. Required for the process-exit
    /// poll/force-terminate step; if not supplied, that step is skipped entirely (the process
    /// may leak — captured only best-effort at batch startup).
    /// </param>
    /// <param name="logger">Logger for diagnostic output (optional).</param>
    /// <param name="filePath">File path for diagnostic messages (optional).</param>
    public static void CloseAndQuit(
        PowerPoint.Presentation? presentation,
        PowerPoint.Application? app,
        int? processId,
        ILogger? logger = null,
        string? filePath = null)
    {
        logger ??= NullLogger.Instance;
        string fileName = string.IsNullOrEmpty(filePath) ? "unknown" : Path.GetFileName(filePath);

        ClosePresentation(presentation, fileName, logger);
        ComUtilities.Release(ref presentation);

        if (app != null)
        {
            QuitApplication(app, fileName, logger);
            ComUtilities.Release(ref app);
        }

        // Force the CLR to process any deferred RCW finalization before we check whether the
        // PowerPoint process has actually exited. Without this, the process can linger until the
        // runtime notices the COM references are gone.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        if (processId.HasValue)
        {
            WaitForProcessExitOrEscalate(processId.Value, fileName, logger);
        }
    }

    private static void ClosePresentation(PowerPoint.Presentation? presentation, string fileName, ILogger logger)
    {
        if (presentation == null) return;

        for (int attempt = 1; attempt <= CloseMaxAttempts; attempt++)
        {
            try
            {
                presentation.Close();
                LogDebugSafe(logger, "Presentation {FileName} closed successfully (attempt {Attempt})", fileName, attempt);
                return;
            }
            catch (COMException ex) when (
                attempt < CloseMaxAttempts &&
                (ex.HResult == ResiliencePipelines.RpcServerCallRetryLater ||
                 ex.HResult == ResiliencePipelines.RpcCallRejected))
            {
                int delayMs = CloseRetryDelayMs * attempt;
                LogDebugSafe(
                    logger,
                    "Presentation close attempt {Attempt} for {FileName} got transient COM busy (0x{HResult:X8}), retrying in {Delay}ms",
                    attempt, fileName, ex.HResult, delayMs);
                Thread.Sleep(delayMs);
            }
            catch (Exception ex)
            {
                // Non-transient failure (or retries exhausted) — log and continue with cleanup;
                // this is a best-effort shutdown step, not something callers can act on.
                logger.LogWarning(ex, "Failed to close presentation {FileName} cleanly during shutdown - continuing with cleanup", fileName);
                return;
            }
        }
    }

    private static void QuitApplication(PowerPoint.Application app, string fileName, ILogger logger)
    {
        var stopwatch = Stopwatch.StartNew();
        int attemptNumber = 0;

        try
        {
            QuitPipeline.Execute(() =>
            {
                attemptNumber++;
                try
                {
                    LogDebugSafe(logger, "Quit attempt {Attempt} for {FileName}", attemptNumber, fileName);
                    app.Quit();
                    LogDebugSafe(
                        logger,
                        "Quit attempt {Attempt} succeeded for {FileName} after {Elapsed}ms",
                        attemptNumber, fileName, stopwatch.ElapsedMilliseconds);
                }
                catch (COMException ex)
                {
                    logger.LogWarning(
                        ex,
                        "Quit attempt {Attempt} failed for {FileName} (HResult: 0x{HResult:X8})",
                        attemptNumber, fileName, ex.HResult);
                    throw; // Let the pipeline decide whether to retry.
                }
            });
        }
        catch (COMException ex)
        {
            // Retries exhausted, or a fatal RPC error (process already gone/disconnected).
            // Proceed with cleanup regardless — the process-exit poll below decides whether
            // force-termination is actually needed.
            logger.LogWarning(
                ex,
                "PowerPoint.Quit() did not succeed for {FileName} after {Attempts} attempt(s) (HResult: 0x{HResult:X8}) in {Elapsed}ms - proceeding with cleanup",
                fileName, attemptNumber, ex.HResult, stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            logger.LogWarning(
                ex,
                "Unexpected error quitting PowerPoint for {FileName} after {Attempts} attempt(s) - proceeding with cleanup",
                fileName, attemptNumber);
        }
    }

    private static void WaitForProcessExitOrEscalate(int processId, string fileName, ILogger logger)
    {
        var stopwatch = Stopwatch.StartNew();
        var delay = TimeSpan.FromMilliseconds(InitialPollDelayMs);

        while (stopwatch.Elapsed < ComInteropConstants.PowerPointProcessExitGracePeriod)
        {
            if (!IsProcessAlive(processId))
            {
                LogDebugSafe(
                    logger,
                    "PowerPoint process {ProcessId} for {FileName} exited normally after {Elapsed}ms - no force-termination needed",
                    processId, fileName, stopwatch.ElapsedMilliseconds);
                return;
            }

            // KernelSleep (not Thread.Sleep) so this poll genuinely waits the full backoff
            // interval regardless of any COM callbacks — this runs after the STA thread's
            // message pump has already exited, so there is nothing left to service anyway.
            ComUtilities.KernelSleep((int)delay.TotalMilliseconds);
            delay = TimeSpan.FromMilliseconds(Math.Min(delay.TotalMilliseconds * 2, MaxPollDelayMs));
        }

        if (!IsProcessAlive(processId))
        {
            LogDebugSafe(
                logger,
                "PowerPoint process {ProcessId} for {FileName} exited right at the grace-period boundary - no force-termination needed",
                processId, fileName);
            return;
        }

        // Grace period exhausted and the process is still alive: this exceeds the known benign
        // ~90-100s Office-cleanup lingering window, so treat it as a genuine hang and
        // force-terminate as a last resort.
        logger.LogWarning(
            "PowerPoint process {ProcessId} for {FileName} did not exit within the {GraceSeconds}s grace period after Quit(). " +
            "Treating as a hung process and force-terminating.",
            processId, fileName, ComInteropConstants.PowerPointProcessExitGracePeriod.TotalSeconds);

        TryKillProcess(processId, fileName, logger);
    }

    private static bool IsProcessAlive(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static void TryKillProcess(int processId, string fileName, ILogger logger)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.HasExited)
            {
                process.Kill();
                process.WaitForExit(5000);
                logger.LogWarning("Force-terminated hung PowerPoint process {ProcessId} for {FileName}", processId, fileName);
            }
        }
        catch (ArgumentException)
        {
            // Already exited between the check and the kill attempt - nothing to do.
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to force-terminate PowerPoint process {ProcessId} for {FileName} - process may leak", processId, fileName);
        }
    }

    /// <summary>
    /// Logs at Debug level only if Debug is actually enabled — avoids CA1873 (unconditional
    /// argument evaluation for a log level that's typically disabled in production).
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Usage",
        "CA2254:Template should be a static expression",
        Justification = "Generic pass-through Debug-logging helper; every call site below passes a compile-time string literal template.")]
    private static void LogDebugSafe(ILogger logger, string message, params object?[] args)
    {
        if (logger.IsEnabled(LogLevel.Debug))
        {
            logger.LogDebug(message, args);
        }
    }
}
