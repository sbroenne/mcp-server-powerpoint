// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using System.IO.Pipelines;
using ModelContextProtocol.Client;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Focused regression test for the create-and-open fix
/// (.squad/decisions/inbox/brett-create-and-open.md): <c>create_presentation</c> must return a
/// sessionId promptly instead of blocking on <c>IPresentationBatch.Dispose()</c>'s
/// grace-period/force-kill sequence (previously ~90-210s, see
/// .squad/decisions/inbox/ripley-create-presentation-blocks-on-dispose.md). Kept in its own file
/// and its own PowerPoint instance so the timing assertion is not skewed by any other test's COM
/// activity in the shared "ProgramTransport" collection.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpCreatePresentationLatencyTests : IAsyncLifetime, IAsyncDisposable
{
    /// <summary>
    /// Generous upper bound for the create_presentation call itself — well below the old
    /// ~90-210s Dispose()-blocking behavior, but with headroom for real (cold-start) PowerPoint
    /// COM activation time, which is unrelated to the fix and observed to take up to ~20s on a
    /// loaded dev box.
    /// </summary>
    private static readonly TimeSpan CreateCallBudget = TimeSpan.FromSeconds(45);

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testPresentationFile;

    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpCreatePresentationLatencyTests(ITestOutputHelper output)
    {
        _output = output;

        _tempDir = Path.Join(Path.GetTempPath(), $"McpCreateLatencyTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testPresentationFile = Path.Join(_tempDir, "CreateLatency.pptx");

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "CreateLatencyTestClient",
            _cts.Token);

        _output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");
    }

    public async Task DisposeAsync() => await DisposeAsyncCore();

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        // NOTE: server-host shutdown below may still take a while (it awaits the background
        // dispose of the session opened by create_presentation) — that is expected and is NOT
        // part of what this test measures; only the create_presentation call itself must be fast.
        await ProgramTransportTestHost.StopAsync(
            _client,
            _clientToServerPipe,
            _serverToClientPipe,
            _serverTask,
            _output);

        _cts.Dispose();

        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors — best effort only.
            }
        }
    }

    /// <summary>
    /// create_presentation must return success + a non-empty sessionId + a real file on disk in
    /// well under the old ~90-210s Dispose()-blocking window, because it no longer disposes the
    /// batch synchronously — it leaves the session open (create-and-open semantics).
    /// </summary>
    [Fact]
    public async Task CreatePresentation_ReturnsSessionIdQuickly_DoesNotBlockOnDispose()
    {
        var stopwatch = Stopwatch.StartNew();

        var createResult = await McpToolCallHelper.CallToolAsync(
            _client!,
            "create_presentation",
            new Dictionary<string, object?> { ["filePath"] = _testPresentationFile },
            _cts.Token);

        stopwatch.Stop();

        McpToolCallHelper.AssertSuccess(createResult, "create_presentation");

        var sessionId = McpToolCallHelper.GetString(createResult, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId), $"Expected a non-empty sessionId in response: {createResult}");
        Assert.True(File.Exists(_testPresentationFile), $"Expected file to exist: {_testPresentationFile}");
        Assert.Equal(_testPresentationFile, McpToolCallHelper.GetString(createResult, "presentationPath"));

        _output.WriteLine($"✓ create_presentation returned sessionId={sessionId} in {stopwatch.Elapsed.TotalSeconds:N1}s");

        Assert.True(
            stopwatch.Elapsed < CreateCallBudget,
            $"create_presentation took {stopwatch.Elapsed.TotalSeconds:N1}s, expected under {CreateCallBudget.TotalSeconds:N0}s " +
            "(the whole point of the create-and-open fix is that it no longer blocks on batch Dispose()).");

        // Clean up the session we opened. close_presentation is itself async (removes from the
        // registry and starts disposal on a background task) — this call should also be fast;
        // the actual PowerPoint shutdown happens after this test returns, during host teardown.
        var closeResult = await McpToolCallHelper.CallToolAsync(
            _client!,
            "close_presentation",
            new Dictionary<string, object?> { ["sessionId"] = sessionId },
            _cts.Token);
        McpToolCallHelper.AssertSuccess(closeResult, "close_presentation");
        _output.WriteLine("✓ close_presentation succeeded (background dispose started)");
    }
}
