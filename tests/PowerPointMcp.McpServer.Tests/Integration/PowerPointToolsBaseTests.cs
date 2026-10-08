using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.PowerPointMcp.McpServer.Tools;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpBoundary")]
public sealed class PowerPointToolsBaseTests
{
    [Fact]
    public async Task ExecuteToolActionAsync_CancelsWhileAsyncOperationIsStillRunning()
    {
        using var cancellation = new CancellationTokenSource();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        var call = PowerPointToolsBase.ExecuteToolActionAsync("slide", "get-count", () =>
        {
            started.SetResult();
            return completion.Task;
        }, cancellation.Token);

        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => call.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.False(completion.Task.IsCompleted);
        }
        finally
        {
            completion.TrySetResult("""{"success":true}""");
        }
    }

    [Fact]
    public async Task ExecuteToolActionAsync_PropagatesRequestedCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PowerPointToolsBase.ExecuteToolActionAsync(
                "slide",
                () => throw new InvalidOperationException("Operation must not run."),
                cancellation.Token));
    }

    [Fact]
    public async Task ExecuteToolActionAsync_PropagatesRequestedCancellationFromAsyncOperation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PowerPointToolsBase.ExecuteToolActionAsync(
                "slide",
                "list",
                () => Task.FromResult("{}"),
                cancellation.Token));
    }

    [Fact]
    public void StdioConsoleFormatter_SuppressesExpectedSdkCancellationWarning()
    {
        var formatter = new StdioConsoleFormatter();
        var entry = new LogEntry<string>(
            LogLevel.Warning,
            "ModelContextProtocol.Server.McpServer",
            new EventId(1),
            "Request cancelled.",
            new OperationCanceledException(),
            static (state, _) => state);
        using var writer = new StringWriter();

        formatter.Write(entry, null, writer);

        Assert.Equal(string.Empty, writer.ToString());
    }
}
