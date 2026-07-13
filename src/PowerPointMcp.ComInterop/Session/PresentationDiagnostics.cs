namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Diagnostic tracing for PowerPoint COM session operations.
/// Ported from mcp-server-excel's ExcelMcp.ComInterop.Session.SessionDiagnostics.
/// </summary>
public static class PresentationDiagnostics
{
    private const string DiagnosticsEnvVar = "POWERPOINTMCP_DIAGNOSTICS";

    /// <summary>
    /// Gets whether diagnostic tracing is enabled via the POWERPOINTMCP_DIAGNOSTICS environment variable.
    /// </summary>
    public static bool IsEnabled =>
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "1", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Writes a diagnostic message to stderr if diagnostics are enabled.
    /// </summary>
    /// <param name="message">Message to write</param>
    public static void WriteStdErr(string message)
    {
        if (IsEnabled)
        {
            Console.Error.WriteLine(message);
        }
    }
}
