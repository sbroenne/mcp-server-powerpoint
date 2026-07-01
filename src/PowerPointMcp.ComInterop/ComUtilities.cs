using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop;

/// <summary>
/// Low-level COM interop utilities for PowerPoint automation.
/// Provides helpers for COM object lifecycle management.
/// </summary>
/// <remarks>
/// Ported from mcp-server-excel's ComUtilities. Only the Office-COM-generic helpers
/// (Release, KernelSleep) were carried over verbatim; the Excel-specific "Find*" helpers
/// (FindQuery, FindName, FindSheet, FindConnection) were dropped and will be replaced with
/// PowerPoint-specific equivalents (e.g. FindSlide, FindShape) as those Core command domains
/// are implemented.
/// </remarks>
public static class ComUtilities
{
    /// <summary>
    /// Safely releases a COM object and sets the reference to null.
    /// </summary>
    /// <remarks>
    /// Use this helper to release intermediate COM objects (like slides, shapes, text frames)
    /// to prevent the PowerPoint process from staying open. This is especially important when
    /// iterating through collections or accessing multiple COM properties.
    /// </remarks>
    public static void Release<T>(ref T? comObject) where T : class
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception)
            {
                // Ignore errors during release — COM object may already be released or RPC disconnected
            }
            comObject = null;
        }
    }

    /// <summary>
    /// Safely attempts to quit a PowerPoint application COM object.
    /// This is a fire-and-forget cleanup helper - errors are swallowed.
    /// </summary>
    /// <remarks>
    /// Use this for cleanup scenarios where you want to quit PowerPoint but don't
    /// need to handle or report errors. For production shutdown with retry logic,
    /// a PresentationShutdownService (mirroring ExcelShutdownService) should be added
    /// alongside the full batch implementation.
    /// </remarks>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("CodeQuality", "CS8602", Justification = "Dynamic COM interop - Quit exists on PowerPoint.Application")]
    public static void TryQuitPowerPoint(PowerPoint.Application? powerPoint)
    {
        if (powerPoint == null) return;

        try
        {
            powerPoint.Quit();
        }
        catch (Exception)
        {
            // Swallow errors during cleanup — PowerPoint may already be gone
        }
    }

    /// <summary>
    /// Safely gets a string property from a COM object, returning empty string if null.
    /// </summary>
    public static string SafeGetString(dynamic? obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "Name" => obj.Name,
                _ => null
            };
            return value?.ToString() ?? string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Safely gets an integer property from a COM object, returning 0 if null or invalid.
    /// </summary>
    public static int SafeGetInt(dynamic? obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "Count" => obj.Count,
                _ => 0
            };
            return Convert.ToInt32(value);
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [DllImport("kernel32.dll")]
    private static extern void Sleep(uint dwMilliseconds);

    /// <summary>
    /// Kernel-level sleep that does NOT pump the STA COM message queue.
    /// Unlike Thread.Sleep (which wakes early on every incoming COM event), this calls
    /// Win32 Sleep() directly — the thread genuinely sleeps for the full interval regardless
    /// of COM callbacks. Safe to use when polling for completion of an operation driven by
    /// PowerPoint's own internals, where our polling thread does not need to service callbacks.
    /// </summary>
    public static void KernelSleep(int milliseconds) =>
        Sleep((uint)Math.Max(0, milliseconds));
}
