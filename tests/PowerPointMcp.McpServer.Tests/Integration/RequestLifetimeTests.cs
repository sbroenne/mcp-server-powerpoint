using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.McpServer.Infrastructure;
using Sbroenne.PowerPointMcp.McpServer.Tools;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresPowerPoint", "true")]
public sealed class RequestLifetimeTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ClipboardWait_ReportsWholeRequestAndRetainsLateFailure(bool usePipe)
    {
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-request-{Guid.NewGuid():N}.pptx");
        string pipeName = $"PowerPointMcp_Request_{Guid.NewGuid():N}";
        using var service = new PowerPointMcpService();
        string sessionId = service.Sessions.Create(path);
        Assert.True(service.Sessions.TryGet(sessionId, out var batch));
        Assert.True(new SlideCommands().AddBlank(batch).Success);
        Assert.True(new ShapeCommands().AddRectangle(batch, 1, 10, 10, 40, 40).Success);
        var server = usePipe ? service.RunAsync(pipeName) : Task.CompletedTask;
        using var acquired = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var holder = Task.Run(() =>
        {
            using var mutex = new Mutex(ShapeCommands.ClipboardMutexName, ShapeCommands.ClipboardMutexOptions);
            bool held;
            try { held = mutex.WaitOne(TimeSpan.FromSeconds(10)); }
            catch (AbandonedMutexException) { held = true; }
            Assert.True(held);
            acquired.Set();
            try { Assert.True(release.Wait(TimeSpan.FromSeconds(90))); }
            finally { mutex.ReleaseMutex(); }
        });
        using var cancellation = new CancellationTokenSource();
        Task? call = null;
        try
        {
            Assert.True(acquired.Wait(TimeSpan.FromSeconds(15)));
            var before = batch.GetOperationStatus().StartedAtUtc;
            object args = new { slideIndex = 1, shapeIndex = 1, targetSlideIndex = 2 };
            call = usePipe
                ? SendPipeRequestAsync()
                : PowerPointToolsBase.ExecuteToolActionAsync("shape", "copy-to-slide",
                    () => ServiceBridge.ForwardToServiceAsync(service, "shape.copy-to-slide",
                        sessionId, args, cancellation.Token), cancellation.Token, service.Sessions, sessionId);

            using var startedTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (batch.GetOperationStatus().StartedAtUtc == before)
                await Task.Delay(20, startedTimeout.Token);
            await Task.Delay(200);
            Assert.False(call.IsCompleted);
            var pending = Assert.Single(service.Sessions.List()).Operation!;
            Assert.True(pending.ExecutionPending, "The request is still waiting for the clipboard after its validation callback.");
            Assert.Equal("running", pending.State);
            Assert.Throws<InvalidOperationException>(() => batch.Execute((context, token) => 0));

            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => call.WaitAsync(TimeSpan.FromSeconds(2)));
            using var cancellationTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (!batch.GetOperationStatus().WaitExpired)
                await Task.Delay(20, cancellationTimeout.Token);
            Assert.True(batch.GetOperationStatus().ExecutionPending);

            using var completionTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(70));
            while (batch.GetOperationStatus().ExecutionPending)
                await Task.Delay(50, completionTimeout.Token);
            var failed = batch.GetOperationStatus();
            Assert.Equal("failed", failed.State);
            Assert.True(failed.WaitExpired);
            Assert.Contains("clipboard", failed.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, new ShapeCommands().GetCount(batch, 2).ShapeCount);

            async Task SendPipeRequestAsync()
            {
                using var client = new ServiceClient(pipeName, requestTimeout: TimeSpan.FromSeconds(80));
                await client.SendAsync(new ServiceRequest
                {
                    Command = "shape.copy-to-slide",
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
                }, cancellation.Token);
            }
        }
        finally
        {
            release.Set();
            await holder.WaitAsync(TimeSpan.FromSeconds(5));
            if (call != null)
            {
                try { await call.WaitAsync(TimeSpan.FromSeconds(5)); }
                catch (OperationCanceledException) { }
            }
            using var drained = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (batch.GetOperationStatus().ExecutionPending)
                await Task.Delay(20, drained.Token);
            service.Sessions.Close(sessionId);
            service.Sessions.DisposeAll();
            service.RequestShutdown();
            await server.WaitAsync(TimeSpan.FromSeconds(5));
            File.Delete(path);
        }
    }
}
