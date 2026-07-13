using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Shared helpers for PowerPoint MCP tools: JSON serialization options, a tool-boundary
/// execution wrapper, and consistent error-payload formatting.
/// </summary>
/// <remarks>
/// Lean MVP port of mcp-server-excel's <c>ExcelToolsBase</c> (telemetry and service-bridge
/// forwarding intentionally omitted — see .squad/decisions.md, Dallas's architecture pass).
///
/// Rule 1/1b boundary: Core commands return <c>{Domain}OperationResult</c> with a
/// Success/ErrorMessage invariant. Expected bad input already surfaces as Success=false and is
/// serialized as an error payload — never thrown. Unexpected COM exceptions propagate out of Core
/// and are caught ONLY here, at the tool boundary, then serialized into a structured error so the
/// MCP host never crashes.
/// </remarks>
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)]
public static class PowerPointToolsBase
{
    /// <summary>
    /// JSON options tuned for LLM token efficiency: compact output, camelCase names, null
    /// properties omitted, and string (rather than numeric) enum values.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Serializes an arbitrary payload with the shared <see cref="JsonOptions"/>.
    /// </summary>
    public static string Serialize(object payload) => JsonSerializer.Serialize(payload, JsonOptions);

    /// <summary>
    /// Executes a tool operation at the MCP boundary. Expected failures are expected to already be
    /// encoded as Success=false payloads by the operation; any unexpected exception is logged to
    /// stderr and serialized into a structured error so the host stays alive (Rule 1b).
    /// </summary>
    /// <param name="toolName">Tool name for error context (e.g. "presentation").</param>
    /// <param name="operation">The synchronous operation producing a JSON response string.</param>
    /// <returns>The operation's JSON response, or a serialized error payload on exception.</returns>
    public static string ExecuteToolAction(string toolName, Func<string> operation)
    {
        try
        {
            return operation();
        }
#pragma warning disable CA1031 // Top-of-tool handler: unexpected exceptions must be serialized, not crash the MCP host.
        catch (Exception ex)
        {
            if (ex is COMException comEx)
            {
                Console.Error.WriteLine(
                    $"[PowerPointMcp] COM Exception in {toolName}: HResult=0x{comEx.HResult:X8}, Message={comEx.Message}");
            }
            else
            {
                Console.Error.WriteLine($"[PowerPointMcp] Exception in {toolName}: {ex.GetType().Name}: {ex.Message}");
            }

            return SerializeToolError(toolName, ex);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Overload used by generated action-dispatch tools: includes the resolved action string
    /// (e.g. "add-chart") in stderr diagnostics for easier troubleshooting, since a single
    /// generated tool (e.g. "chart") now covers many actions.
    /// </summary>
    /// <param name="toolName">Tool name for error context (e.g. "chart").</param>
    /// <param name="actionName">Kebab-case action name for error context (e.g. "add-chart").</param>
    /// <param name="operation">The synchronous operation producing a JSON response string.</param>
    /// <returns>The operation's JSON response, or a serialized error payload on exception.</returns>
    public static string ExecuteToolAction(string toolName, string actionName, Func<string> operation)
    {
        var context = $"{toolName}.{actionName}";
        try
        {
            return operation();
        }
#pragma warning disable CA1031 // Top-of-tool handler: unexpected exceptions must be serialized, not crash the MCP host.
        catch (Exception ex)
        {
            if (ex is COMException comEx)
            {
                Console.Error.WriteLine(
                    $"[PowerPointMcp] COM Exception in {context}: HResult=0x{comEx.HResult:X8}, Message={comEx.Message}");
            }
            else
            {
                Console.Error.WriteLine($"[PowerPointMcp] Exception in {context}: {ex.GetType().Name}: {ex.Message}");
            }

            return SerializeToolError(context, ex);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Serializes an exception into a consistent error payload
    /// (<c>success=false</c>, <c>isError=true</c>, plus COM diagnostics where available).
    /// </summary>
    public static string SerializeToolError(string toolName, Exception ex)
    {
        var errorMessage = $"{toolName} failed: {ex.Message}";
        string exceptionType = ex.GetType().Name;
        string? hresult = null;
        string? innerError = null;

        if (ex is COMException comEx)
        {
            hresult = $"0x{comEx.HResult:X8}";
            errorMessage += $" [COM Error: {hresult}]";
        }

        if (ex.InnerException != null)
        {
            innerError = ex.InnerException.Message;
            if (ex.InnerException is COMException innerComEx)
            {
                innerError += $" [COM: 0x{innerComEx.HResult:X8}]";
            }
        }

        return Serialize(new
        {
            success = false,
            errorMessage,
            exceptionType,
            hresult,
            innerError,
            isError = true
        });
    }

    /// <summary>
    /// Serializes a structured "bad input" error (Success=false) without throwing. Use for
    /// expected, caller-correctable validation failures (missing path, unknown session, etc.).
    /// </summary>
    public static string ValidationError(string message) => Serialize(new
    {
        success = false,
        errorMessage = message,
        isError = true
    });
}
