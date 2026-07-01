using Microsoft.Extensions.Logging;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Minimal close/quit helper for PowerPoint sessions.
/// </summary>
/// <remarks>
/// This is a first-pass, non-resilient version. mcp-server-excel's ExcelShutdownService has
/// exponential-backoff retry for transient COM busy errors on Close()/Quit() — that hardening
/// should be ported here (as PresentationShutdownService) once the basic vertical slice is
/// validated against real PowerPoint COM behavior.
/// </remarks>
internal static class PresentationShutdownService
{
    public static void CloseAndQuit(PowerPoint.Presentation? presentation, PowerPoint.Application? app, ILogger logger)
    {
        try
        {
            presentation?.Close();
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Failed to close presentation cleanly during shutdown.");
        }

        ComUtilities.Release(ref presentation);

        if (app != null)
        {
            ComUtilities.TryQuitPowerPoint(app);
            ComUtilities.Release(ref app);
        }
    }
}
