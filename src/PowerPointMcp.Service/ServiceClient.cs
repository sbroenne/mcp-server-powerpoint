using System.IO.Pipes;
using Sbroenne.PowerPointMcp.Service.Rpc;
using StreamJsonRpc;

namespace Sbroenne.PowerPointMcp.Service;

/// <summary>
/// Client for communicating with the PowerPointMCP CLI daemon via named pipe + StreamJsonRpc.
/// Each call creates a new pipe connection, makes one RPC call, and disconnects. Ported from
/// mcp-server-excel's ExcelMcp.Service.ServiceClient.
/// </summary>
public sealed class ServiceClient : IDisposable
{
    private readonly string _pipeName;
    private readonly TimeSpan _connectTimeout;
    private readonly TimeSpan _requestTimeout;
    private bool _disposed;

    /// <summary>Default timeout for establishing the named pipe connection.</summary>
    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);

    /// <summary>Default timeout for the whole RPC round-trip (long enough that any caller-supplied timeout wins first).</summary>
    public static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromHours(2);

    /// <summary>Creates a client bound to a specific daemon pipe name.</summary>
    /// <param name="pipeName">The named pipe to connect to.</param>
    /// <param name="connectTimeout">Optional override for the pipe connect timeout.</param>
    /// <param name="requestTimeout">Optional override for the overall request timeout.</param>
    public ServiceClient(string pipeName, TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
        _pipeName = pipeName;
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
    }

    /// <summary>Sends a request to the daemon and waits for its response via StreamJsonRpc.</summary>
    public async Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        using var pipe = ServiceSecurity.CreateClient(_pipeName);
        using var connectCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        connectCts.CancelAfter(_connectTimeout);

        try
        {
            await pipe.ConnectAsync((int)_connectTimeout.TotalMilliseconds, connectCts.Token);

            var proxy = JsonRpc.Attach<IPowerPointDaemonRpc>(pipe);
            try
            {
                using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                using var disconnectMonitorCts = CancellationTokenSource.CreateLinkedTokenSource(requestCts.Token);
                requestCts.CancelAfter(_requestTimeout);

                var callTask = proxy.ProcessCommandAsync(request);
                var disconnectTask = WaitForPipeDisconnectAsync(pipe, disconnectMonitorCts.Token);
                var completed = await Task.WhenAny(callTask, disconnectTask);
                if (completed == disconnectTask
                    && disconnectTask.IsCompletedSuccessfully
                    && disconnectTask.Result
                    && !callTask.IsCompleted)
                {
                    return CreateConnectionLostResponse(request);
                }

                await disconnectMonitorCts.CancelAsync();
                return await callTask.WaitAsync(requestCts.Token);
            }
            finally
            {
                ((IDisposable)proxy).Dispose();
            }
        }
        catch (TimeoutException)
        {
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "Timeout",
                ErrorMessage = "Daemon connection timed out",
                ExceptionType = nameof(TimeoutException)
            };
        }
        catch (OperationCanceledException) when (connectCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "Timeout",
                ErrorMessage = "Daemon connection timed out",
                ExceptionType = nameof(OperationCanceledException)
            };
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "Timeout",
                ErrorMessage = "Daemon request timed out",
                ExceptionType = nameof(OperationCanceledException)
            };
        }
        catch (ConnectionLostException)
        {
            return CreateConnectionLostResponse(request);
        }
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = "Cannot connect to daemon. Is it running?",
                ExceptionType = nameof(IOException)
            };
        }
    }

    private static async Task<bool> WaitForPipeDisconnectAsync(Stream pipe, CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                if (pipe is PipeStream pipeStream && !pipeStream.IsConnected)
                {
                    return true;
                }

                await Task.Delay(TimeSpan.FromMilliseconds(250), cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Caller completed or request timed out; this is not a disconnect signal.
        }
        catch (ObjectDisposedException)
        {
            return true;
        }

        return false;
    }

    private static ServiceResponse CreateConnectionLostResponse(ServiceRequest request)
    {
        return new ServiceResponse
        {
            Success = false,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = "ServiceUnavailable",
            ErrorMessage = "Connection to the daemon was lost while waiting for a response. The daemon may have " +
                "exited or restarted; run 'pptcli service status' or 'pptcli service stop' and retry.",
            ExceptionType = nameof(ConnectionLostException)
        };
    }

    /// <summary>Pings the daemon to check if it's alive.</summary>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await SendAsync(new ServiceRequest { Command = "service.ping" }, cancellationToken);
            return response.Success;
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <summary>Marks this client as disposed; safe to call even though there is no unmanaged state.</summary>
    public void Dispose()
    {
        _disposed = true;
    }
}
