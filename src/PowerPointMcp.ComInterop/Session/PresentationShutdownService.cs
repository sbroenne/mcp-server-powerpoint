extern alias OfficeInterop;

using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Polly;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Centralized service for PowerPoint presentation close and application quit operations.
/// Implements resilient shutdown with retry/backoff for COM busy conditions, ported from
/// mcp-server-excel's <c>ExcelShutdownService</c>.
/// </summary>
/// <remarks>
/// Runs only on the batch's PowerPoint thread. Host shutdown retains unsaved presentations;
/// explicit close discards changes only after any requested save. Neither path terminates
/// PowerPoint, and Quit is allowed only when no presentations remain open.
/// </remarks>
internal static class PresentationShutdownService
{
    private static readonly ResiliencePipeline QuitPipeline = ResiliencePipelines.CreatePowerPointQuitPipeline();

    private const int CloseMaxAttempts = 3;
    private const int CloseRetryDelayMs = 200;

    /// <summary>
    /// Releases presentation and application references, retaining unsaved work on host shutdown.
    /// </summary>
    /// <param name="presentation">Presentation COM object (can be null).</param>
    /// <param name="app">Application COM object (can be null).</param>
    /// <param name="logger">Logger for diagnostic output (optional).</param>
    /// <param name="filePath">File path for diagnostic messages (optional).</param>
    public static void CloseAndQuit(
        PowerPoint.Presentation? presentation,
        PowerPoint.Application? app,
        ILogger? logger = null,
        string? filePath = null,
        Action? afterPresentationClose = null,
        bool preserveUnsavedChanges = false,
        bool leavePresentationOpen = false)
    {
        logger ??= NullLogger.Instance;
        string fileName = string.IsNullOrEmpty(filePath) ? "unknown" : Path.GetFileName(filePath);

        bool retain = leavePresentationOpen;
        if (preserveUnsavedChanges && presentation != null && !retain)
        {
            try
            {
                retain = presentation.Saved != Office.MsoTriState.msoTrue;
            }
            catch (COMException ex)
            {
                logger.LogWarning(ex, "Could not read presentation save state; leaving it open without saving or closing it.");
                retain = true;
            }
        }

        if (!retain)
            ClosePresentation(presentation, fileName, logger);
        afterPresentationClose?.Invoke();
        ComUtilities.Release(ref presentation);

        if (app != null)
        {
            if (!retain && CanQuitApplication(app, logger))
                QuitApplication(app, fileName, logger);
            else
            {
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
                    app.Visible = Office.MsoTriState.msoTrue;
                }
                catch (COMException ex)
                {
                    logger.LogWarning(ex, "PowerPoint was left open, but its window or alerts could not be restored.");
                }
            }
            ComUtilities.Release(ref app);
        }

        // Release deferred COM wrappers without waiting for the Office process to exit.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

    }

    private static bool CanQuitApplication(PowerPoint.Application app, ILogger logger)
    {
        PowerPoint.Presentations? presentations = null;
        try
        {
            presentations = app.Presentations;
            if (presentations.Count == 0)
                return true;
            LogDebugSafe(logger, "Leaving the shared PowerPoint application running because other presentations are still open.");
            return false;
        }
        catch (COMException ex)
        {
            logger.LogWarning(ex, "Cannot confirm that PowerPoint has no other presentations; refusing to quit or terminate the shared application.");
            return false;
        }
        finally
        {
            ComUtilities.Release(ref presentations);
        }
    }

    private static void ClosePresentation(PowerPoint.Presentation? presentation, string fileName, ILogger logger)
    {
        if (presentation == null) return;

        for (int attempt = 1; attempt <= CloseMaxAttempts; attempt++)
        {
            try
            {
                presentation.Saved = Office.MsoTriState.msoTrue;
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
            // Release automation without terminating PowerPoint.
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
