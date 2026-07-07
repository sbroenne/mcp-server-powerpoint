namespace Sbroenne.PowerPointMcp.Service.Rpc;

/// <summary>
/// Typed RPC interface for CLI &lt;-&gt; daemon communication over named pipes, using StreamJsonRpc
/// (JSON-RPC 2.0) instead of a hand-rolled newline-delimited protocol.
/// </summary>
/// <remarks>
/// Ported from mcp-server-excel's <c>IExcelDaemonRpc</c>. Deliberately omits Excel's
/// <c>[JsonRpcContract]</c>/<c>[GenerateShape]</c> (PolyType) attributes — those opt into
/// NativeAOT-safe source-generated proxies, which this project does not need (it is not
/// published as NativeAOT). <c>StreamJsonRpc.JsonRpc.Attach&lt;T&gt;</c> builds a classic
/// reflection-based dynamic proxy for a plain interface with no extra package dependency.
/// </remarks>
public interface IPowerPointDaemonRpc
{
    /// <summary>
    /// Sends a command to the daemon for execution. Wraps
    /// <see cref="PowerPointMcpService.ProcessAsync"/> over the pipe transport.
    /// </summary>
    /// <param name="request">The service request with command, sessionId, and args.</param>
    /// <returns>The service response indicating success/failure with optional result data.</returns>
    Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request);
}
