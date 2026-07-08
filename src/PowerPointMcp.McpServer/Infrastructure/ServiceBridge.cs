using System.Text.Json;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.McpServer.Infrastructure;

/// <summary>
/// In-process bridge from generated MCP action-dispatch tools to the shared
/// <see cref="PowerPointMcpService"/> instance — mirroring mcp-server-excel's
/// <c>ServiceBridge</c> (one shared Service class, consumed in-process here with no named pipe,
/// and via named-pipe/StreamJsonRpc by the separate CLI daemon process).
/// </summary>
/// <remarks>
/// Generated tools (see <c>PowerPointMcp.Generators.Mcp</c>) call <see cref="ForwardToService"/>
/// as their <c>System.Func&lt;string, string, object?, string&gt;</c> "forwardToService"
/// delegate passed into <c>ServiceRegistry.{Category}.RouteAction</c>. This keeps the Rule 1/1b
/// error-shape consistent with the rest of the MCP surface: expected failures (bad session id,
/// validation errors surfaced by Core) come back as <c>Success=false</c> on the
/// <see cref="ServiceResponse"/> and are serialized here into the same error JSON shape used by
/// <see cref="Sbroenne.PowerPointMcp.McpServer.Tools.PowerPointToolsBase.SerializeToolError"/>;
/// unexpected exceptions are already caught inside <see cref="PowerPointMcpService.ProcessAsync"/>
/// and returned as a Success=false response too, so this bridge never needs its own try/catch.
/// </remarks>
public static class ServiceBridge
{
    /// <summary>
    /// JSON options for serializing generated tool call arguments before forwarding to
    /// <see cref="PowerPointMcpService.ProcessAsync"/>. Must match
    /// <see cref="ServiceProtocol.JsonOptions"/>'s camelCase naming policy so property names line
    /// up with the generated <c>{Action}Args</c> classes' <c>DeserializeArgs</c> call.
    /// </summary>
    private static readonly JsonSerializerOptions ArgsJsonOptions = ServiceProtocol.JsonOptions;

    /// <summary>
    /// Forwards a generated action-dispatch call to the in-process <see cref="PowerPointMcpService"/>
    /// and returns the JSON result string (or a structured error payload) synchronously.
    /// </summary>
    /// <param name="service">The shared, DI-injected service instance.</param>
    /// <param name="command">Full "category.action" command string (e.g. "chart.add-chart").</param>
    /// <param name="sessionId">The session id supplied by the caller.</param>
    /// <param name="args">The anonymous args object built by the generated <c>RouteAction</c> method.</param>
    public static string ForwardToService(PowerPointMcpService service, string command, string? sessionId, object? args)
    {
        var argsJson = args != null ? JsonSerializer.Serialize(args, ArgsJsonOptions) : null;
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = argsJson,
            Source = "mcp"
        };

        var response = service.ProcessAsync(request).GetAwaiter().GetResult();

        if (response.Success)
        {
            return response.Result ?? JsonSerializer.Serialize(new { success = true }, ArgsJsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = false,
            errorMessage = response.ErrorMessage,
            errorCategory = response.ErrorCategory,
            exceptionType = response.ExceptionType,
            hresult = response.HResult,
            innerError = response.InnerError,
            isError = true
        }, ArgsJsonOptions);
    }
}
