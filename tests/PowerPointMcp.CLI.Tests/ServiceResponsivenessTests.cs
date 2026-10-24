using System.IO.Pipes;
using Sbroenne.PowerPointMcp.Service;
using StreamJsonRpc;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class ServiceResponsivenessTests
{
    [Fact]
    public async Task CallerCancellation_CancelsDisconnectedServiceRequest()
    {
        string pipeName = $"PowerPointMcp_Cancellation_{Guid.NewGuid():N}";
        await using var pipe = ServiceSecurity.CreateSecureServer(pipeName);
        var probe = new CancellationProbe();
        var connected = pipe.WaitForConnectionAsync();
        using var cancellation = new CancellationTokenSource();
        using var client = new ServiceClient(pipeName);
        var request = client.SendAsync(new ServiceRequest { Command = "service.status" }, cancellation.Token);
        await connected.WaitAsync(TimeSpan.FromSeconds(2));
        using var rpc = new JsonRpc(pipe, pipe, probe)
        {
            CancelLocallyInvokedMethodsWhenConnectionIsClosed = true
        };
        rpc.StartListening();
        try
        {
            await probe.Started.Task.WaitAsync(TimeSpan.FromSeconds(2));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
            await Task.WhenAny(probe.Cancelled.Task, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.True(probe.Cancelled.Task.IsCompleted, "The caller's cancellation must reach the service.");
        }
        finally
        {
            probe.Release.TrySetResult();
        }
    }

    private sealed class CancellationProbe
    {
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Cancelled { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request, CancellationToken cancellationToken)
        {
            using var registration = cancellationToken.Register(() => Cancelled.TrySetResult());
            Started.TrySetResult();
            await Release.Task;
            return new ServiceResponse { Success = true, Command = request.Command };
        }
    }

    [Fact]
    public async Task IdleConnections_DoNotBlockStatusOrShutdown()
    {
        string pipeName = $"PowerPointMcp_Pressure_{Guid.NewGuid():N}";
        using var service = new PowerPointMcpService();
        var server = service.RunAsync(pipeName);
        var connections = new List<NamedPipeClientStream>();
        try
        {
            for (int i = 0; i < 12; i++)
            {
                var pipe = ServiceSecurity.CreateClient(pipeName);
                connections.Add(pipe);
                await pipe.ConnectAsync(2000);
            }

            using var client = new ServiceClient(pipeName, requestTimeout: TimeSpan.FromSeconds(2));
            var status = client.SendAsync(new ServiceRequest { Command = "service.status" });
            await Task.WhenAny(status, Task.Delay(TimeSpan.FromSeconds(3)));
            Assert.True(status.IsCompleted, "Idle clients must not block service status.");
            Assert.True((await status).Success);

            service.RequestShutdown();
            await Task.WhenAny(server, Task.Delay(TimeSpan.FromSeconds(3)));
            Assert.True(server.IsCompleted, "Shutdown must not wait for idle clients to disconnect.");
        }
        finally
        {
            foreach (var connection in connections)
                await connection.DisposeAsync();
            service.RequestShutdown();
            await server.WaitAsync(TimeSpan.FromSeconds(10));
        }
    }

    [Theory]
    [InlineData(3000)]
    [InlineData(500)]
    public async Task RequestTimeout_BoundsWaitForSilentServer(int connectTimeoutMs)
    {
        string pipeName = $"PowerPointMcp_Deadline_{Guid.NewGuid():N}";
        await using var server = ServiceSecurity.CreateSecureServer(pipeName);
        var connected = server.WaitForConnectionAsync();
        using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromMilliseconds(connectTimeoutMs),
            requestTimeout: TimeSpan.FromSeconds(1));
        var request = client.SendAsync(new ServiceRequest { Command = "service.status" });
        await connected.WaitAsync(TimeSpan.FromSeconds(2));
        try
        {
            await Task.WhenAny(request, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.True(request.IsCompleted, "The request deadline must include waiting for the first response.");
            var response = await request;
            Assert.False(response.Success);
            Assert.Equal("Timeout", response.ErrorCategory);
            Assert.Contains("PowerPoint work may still be running", response.ErrorMessage, StringComparison.Ordinal);
        }
        finally
        {
            await server.DisposeAsync();
            await request.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }
}
