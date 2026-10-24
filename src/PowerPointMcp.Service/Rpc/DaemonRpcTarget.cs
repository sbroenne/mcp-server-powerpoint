using System.Diagnostics;

namespace Sbroenne.PowerPointMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to
/// <see cref="PowerPointMcpService.ProcessAsync"/>. One instance is attached per pipe connection
/// via <c>JsonRpc.Attach(stream, target)</c>. Ported from mcp-server-excel's DaemonRpcTarget.
/// </summary>
internal sealed class DaemonRpcTarget : IPowerPointDaemonRpc
{
    private readonly PowerPointMcpService _service;
    private readonly object _activityLock = new();
    private long _lastActivityTimestamp = Stopwatch.GetTimestamp();
    private int _activeRequests;
    private bool _idleClosing;

    public DaemonRpcTarget(PowerPointMcpService service)
    {
        _service = service;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        lock (_activityLock)
        {
            if (_idleClosing)
                return new ServiceResponse
                {
                    Success = false,
                    Command = request.Command,
                    SessionId = request.SessionId,
                    ErrorCategory = "ServiceUnavailable",
                    ErrorMessage = "The idle service connection expired. Reconnect and retry; no PowerPoint work was started."
                };
            _activeRequests++;
        }
        try
        {
            return await _service.ProcessAsync(request, cancellationToken);
        }
        finally
        {
            lock (_activityLock)
            {
                _activeRequests--;
                _lastActivityTimestamp = Stopwatch.GetTimestamp();
            }
        }
    }

    internal TimeSpan GetIdleWaitTimeout(TimeSpan idleTimeout)
    {
        lock (_activityLock)
        {
            if (_activeRequests > 0)
                return idleTimeout;
            var remaining = idleTimeout - Stopwatch.GetElapsedTime(_lastActivityTimestamp);
            if (remaining > TimeSpan.Zero)
                return remaining;
            _idleClosing = true;
            return TimeSpan.Zero;
        }
    }
}
