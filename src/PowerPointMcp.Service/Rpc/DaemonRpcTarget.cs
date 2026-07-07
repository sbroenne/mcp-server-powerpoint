namespace Sbroenne.PowerPointMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to
/// <see cref="PowerPointMcpService.ProcessAsync"/>. One instance is attached per pipe connection
/// via <c>JsonRpc.Attach(stream, target)</c>. Ported from mcp-server-excel's DaemonRpcTarget.
/// </summary>
internal sealed class DaemonRpcTarget : IPowerPointDaemonRpc
{
    private readonly PowerPointMcpService _service;

    public DaemonRpcTarget(PowerPointMcpService service)
    {
        _service = service;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
    {
        _service.RecordActivity();
        return await _service.ProcessAsync(request);
    }
}
