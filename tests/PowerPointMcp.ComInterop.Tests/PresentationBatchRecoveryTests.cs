using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.ComInterop.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "PresentationBatch")]
public sealed class PresentationBatchRecoveryTests
{
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    public async Task ExpiredWait_PreservesPresentationAndAllowsWorkAfterCompletion(bool cancellable, bool failsLater, bool trackRequest)
    {
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-recovery-{Guid.NewGuid():N}.pptx");
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        using var batch = PresentationSession.CreateNew(path, operationTimeout: TimeSpan.FromSeconds(5));
        try
        {
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var operation = Task.Run(() =>
            {
                if (trackRequest)
                    batch.BeginRequest();
                try
                {
                    return batch.Execute((context, token) =>
                    {
                        started.SetResult();
                        Assert.True(release.Wait(TimeSpan.FromSeconds(20)));
                        if (failsLater)
                            throw new InvalidOperationException("The delayed PowerPoint operation failed.");
                        return context.Presentation.Name;
                    }, cancellable ? cancellation.Token : CancellationToken.None);
                }
                finally
                {
                    if (trackRequest)
                        batch.CompleteRequest("failed", "The request wait ended.");
                }
            });

            try
            {
                await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
                await Task.WhenAny(operation, Task.Delay(TimeSpan.FromSeconds(8)));
                Assert.True(operation.IsCompleted, "The operation timeout must also bound waits with a caller cancellation token.");
                await Assert.ThrowsAsync<TimeoutException>(() => operation);
                var pending = batch.GetOperationStatus();
                Assert.Equal("running", pending.State);
                Assert.True(pending.ExecutionPending);
                Assert.True(pending.WaitExpired);
                Assert.NotNull(pending.StartedAtUtc);
                Assert.Throws<InvalidOperationException>(() => batch.Execute((context, token) => 0));
            }
            finally
            {
                release.Set();
                try { await operation.WaitAsync(TimeSpan.FromSeconds(10)); }
                catch (TimeoutException) { }
            }

            using var completionTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (batch.GetOperationStatus().ExecutionPending)
                await Task.Delay(20, completionTimeout.Token);

            var completed = batch.GetOperationStatus();
            Assert.Equal(failsLater ? "failed" : "completed", completed.State);
            Assert.Equal(failsLater ? "The delayed PowerPoint operation failed." : null, completed.ErrorMessage);
            Assert.True(completed.WaitExpired);
            Assert.True(batch.IsPowerPointProcessAlive());
            Assert.Equal(Path.GetFileName(path), batch.Execute((context, token) => context.Presentation.Name));
            Assert.False(batch.GetOperationStatus().WaitExpired);
        }
        finally
        {
            release.Set();
            batch.Dispose();
            File.Delete(path);
        }
    }
}
