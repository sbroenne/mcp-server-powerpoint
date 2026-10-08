using System.Text.Json;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.McpServer.Tools;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpCancellationSessionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancelledStartup_ReclaimsLateSessionWithoutClosingAnotherPresentation(bool createAfterCancellation)
    {
        var directory = Path.Join(Path.GetTempPath(), $"McpCancellation_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            using var registry = new PresentationSessionRegistry();
            var unrelated = registry.Create(Path.Join(directory, "Existing.pptx"));
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var created = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            using var cancellation = new CancellationTokenSource();
            var call = PowerPointToolsBase.ExecuteToolActionAsync("presentation", "create", async () =>
            {
                started.SetResult();
                if (createAfterCancellation)
                    await release.Task;
                var result = await PresentationTools.Presentation(
                    PresentationToolAction.Create,
                    filePath: Path.Join(directory, "Late.pptx"),
                    registry: registry);
                var json = Assert.IsType<JsonElement>(result.StructuredContent);
                Assert.True(json.GetProperty("success").GetBoolean());
                created.SetResult(json.GetProperty("presentation_session_id").GetString()!);
                if (!createAfterCancellation)
                    await release.Task;
                return json.GetRawText();
            }, cancellation.Token, registry);

            try
            {
                if (createAfterCancellation)
                    await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
                else
                    await created.Task.WaitAsync(TimeSpan.FromSeconds(60));
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => call.WaitAsync(TimeSpan.FromSeconds(2)));
                release.SetResult();
                var lateSession = await created.Task.WaitAsync(TimeSpan.FromSeconds(60));
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                while (registry.TryGet(lateSession, out _))
                    await Task.Delay(20, timeout.Token);

                Assert.True(registry.TryGet(unrelated, out var batch));
                var commands = new SlideCommands();
                var previousCount = commands.GetCount(batch).SlideCount;
                Assert.True(commands.AddBlank(batch).Success);
                Assert.Equal(previousCount + 1, commands.GetCount(batch).SlideCount);
                Assert.True(File.Exists(Path.Join(directory, "Late.pptx")));
            }
            finally
            {
                release.TrySetResult();
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ClosingOneSession_PreservesOtherPresentationInSharedApplication()
    {
        var directory = Path.Join(Path.GetTempPath(), $"McpCancellation_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            using var registry = new PresentationSessionRegistry();
            var unrelated = registry.Create(Path.Join(directory, "Existing.pptx"));
            var affected = registry.Create(Path.Join(directory, "Affected.pptx"));
            Assert.True(registry.TryGet(unrelated, out var unrelatedBatch));
            Assert.True(registry.TryGet(affected, out var affectedBatch));
            var commands = new SlideCommands();
            var count = commands.GetCount(unrelatedBatch).SlideCount;

            affectedBatch.Dispose();
            Assert.True(registry.Close(affected));

            Assert.True(registry.TryGet(unrelated, out var preserved));
            Assert.True(preserved.IsPowerPointProcessAlive());
            Assert.True(commands.AddBlank(preserved).Success);
            Assert.Equal(count + 1, commands.GetCount(preserved).SlideCount);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void DisposedRegistry_RejectsLateStartup()
    {
        var directory = Path.Join(Path.GetTempPath(), $"McpCancellation_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        using var registry = new PresentationSessionRegistry();
        registry.Dispose();
        try
        {
            Assert.Throws<ObjectDisposedException>(() => registry.Create(Path.Join(directory, "Late.pptx")));
            Assert.Empty(registry.List());
        }
        finally
        {
            registry.DisposeAll();
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task CancelledComWait_ClosesOnlyAffectedSessionAndReturnsBeforeComFinishes()
    {
        var directory = Path.Join(Path.GetTempPath(), $"McpCancellation_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            using var registry = new PresentationSessionRegistry();
            var unrelated = registry.Create(Path.Join(directory, "Existing.pptx"));
            var affected = registry.Create(Path.Join(directory, "Affected.pptx"));
            Assert.True(registry.TryGet(affected, out var affectedBatch));
            using var release = new ManualResetEventSlim();
            using var cancellation = new CancellationTokenSource();
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var finished = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var call = PowerPointToolsBase.ExecuteToolActionAsync("slide", "get-count", () =>
            {
                try
                {
                    affectedBatch.Execute((context, token) =>
                    {
                        started.SetResult();
                        Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                    });
                    return """{"success":true}""";
                }
                finally
                {
                    finished.SetResult();
                }
            }, cancellation.Token, registry, affected);

            try
            {
                await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => call.WaitAsync(TimeSpan.FromSeconds(2)));
                Assert.False(finished.Task.IsCompleted);
                Assert.False(registry.TryGet(affected, out _));
                Assert.True(registry.TryGet(unrelated, out var batch));
                Assert.True(new SlideCommands().GetCount(batch).Success);
            }
            finally
            {
                release.Set();
                await finished.Task.WaitAsync(TimeSpan.FromSeconds(10));
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
