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
    /// Observed allowance for PowerPoint's background cleanup after Quit. Process lifetime alone
    /// is not grounds to terminate PowerPoint, which may still contain unsaved work.
    /// </summary>
    public static readonly TimeSpan PowerPointProcessExitGracePeriod = TimeSpan.FromSeconds(150);

    /// <summary>
    /// Maximum wait for the PowerPoint thread during shutdown. If it remains busy, the host
    /// releases automation without terminating PowerPoint or waiting for its process to exit.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = TimeSpan.FromSeconds(5);

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
