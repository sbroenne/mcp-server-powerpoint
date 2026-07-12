using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop;

/// <summary>
/// Constants for PowerPoint COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>Timeout for PowerPoint.Quit() operation (30 seconds).</summary>
    public static readonly TimeSpan PowerPointQuitTimeout = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Grace period for POWERPNT.exe to exit the OS process list after Application.Quit()
    /// returns, before <see cref="Session.PresentationShutdownService"/> force-terminates it
    /// as a last resort. Sized comfortably above the ~90-100 second benign Office-cleanup
    /// lingering window observed in real-COM MCP round-trip testing (see .squad/decisions.md,
    /// 2026-07-01 McpServer Phase 1 MVP entry) so the happy path never force-kills.
    /// </summary>
    public static readonly TimeSpan PowerPointProcessExitGracePeriod = TimeSpan.FromSeconds(150);

    /// <summary>
    /// Timeout for STA thread join after quit.
    /// CRITICAL: must be &gt;= the full worst-case duration of
    /// <see cref="Session.PresentationShutdownService.CloseAndQuit"/> (Close()/Quit() retries +
    /// <see cref="PowerPointProcessExitGracePeriod"/> + force-kill), not just
    /// <see cref="PowerPointQuitTimeout"/>. <c>PresentationBatch.Dispose()</c> runs
    /// <c>PresentationShutdownService</c> synchronously on the STA thread and joins it with this
    /// timeout — if this were shorter than the shutdown service's own worst-case duration,
    /// Dispose() could return (and the host process could exit) BEFORE the process-exit
    /// grace-period poll / force-kill escalation ever runs, silently defeating the safety net and
    /// leaking POWERPNT.exe. On the happy path (PowerPoint quits within seconds) Join() returns as
    /// soon as the STA thread actually finishes, so this long timeout does not slow down normal
    /// shutdown — it only bounds the worst case.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout =
        PowerPointProcessExitGracePeriod + PowerPointQuitTimeout + TimeSpan.FromSeconds(30);

    /// <summary>Timeout for save operations (5 minutes). Large decks with media may take longer.</summary>
    public static readonly TimeSpan SaveOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Default timeout for individual PowerPoint operations (2 minutes).
    /// Can be overridden when creating a session via timeoutSeconds parameter.
    /// </summary>
    public static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromSeconds(120);

    /// <summary>Maximum wait time for session creation file lock acquisition (5 seconds).</summary>
    public static readonly TimeSpan SessionFileLockTimeout = TimeSpan.FromSeconds(5);

    #endregion

    #region Sleep Intervals

    /// <summary>Delay between file lock acquisition retries (100ms).</summary>
    public const int FileLockRetryDelayMs = 100;

    /// <summary>Delay between session lock acquisition retries (200ms).</summary>
    public const int SessionLockRetryDelayMs = 200;

    #endregion

    #region PowerPoint File Formats

    /// <summary>
    /// PowerPoint Open XML Presentation format code (.pptx).
    /// PpSaveAsFileType.ppSaveAsOpenXMLPresentation = 24
    /// </summary>
    public const PowerPoint.PpSaveAsFileType PpSaveAsOpenXmlPresentation =
        PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation;

    /// <summary>
    /// PowerPoint Open XML Macro-Enabled Presentation format code (.pptm).
    /// PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled = 25
    /// </summary>
    public const PowerPoint.PpSaveAsFileType PpSaveAsOpenXmlPresentationMacroEnabled =
        PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled;

    /// <summary>
    /// PowerPoint Open XML Template format code (.potx).
    /// PpSaveAsFileType.ppSaveAsOpenXMLTemplate = 26
    /// </summary>
    public const PowerPoint.PpSaveAsFileType PpSaveAsOpenXmlTemplate =
        PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLTemplate;

    #endregion
}
