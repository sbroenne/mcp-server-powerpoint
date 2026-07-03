// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using Xunit.Abstractions;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Integration.McpToolCallHelper;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Close/shutdown robustness coverage for Brett's async-close design
/// (.squad/decisions/inbox/brett-async-close.md): double-close, concurrent closes of multiple
/// sessions, and a close-then-immediate-host-shutdown race — all driven through the MCP protocol,
/// with a final real-process check that no POWERPNT.exe spawned by this test survives host
/// shutdown.
/// </summary>
/// <remarks>
/// Manages its own <see cref="ProgramTransportTestHost"/> lifecycle inline (rather than via
/// <see cref="IAsyncLifetime"/>) so the test can control exactly when the host stops relative to
/// the last <c>close_presentation</c> call — the "close-then-host-shutdown" race Brett flagged.
/// Only ONE <c>create_presentation</c> call is made (that call alone currently blocks
/// synchronously for the full PowerPoint quit sequence — see
/// .squad/decisions/inbox/ripley-create-presentation-blocks-on-dispose.md); the resulting file is
/// copied on disk (no COM) to produce independent files for each of the four sessions this test
/// opens, keeping total runtime bounded while still exercising four real PowerPoint processes.
/// </remarks>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ShutdownRobustness")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpShutdownRobustnessTests
{
    private readonly ITestOutputHelper _output;

    public McpShutdownRobustnessTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public async Task CloseRobustness_DoubleCloseConcurrentCloseAndHostShutdown_LeavesNoOrphanedProcesses()
    {
        var baselinePowerPointPids = GetCurrentPowerPointPids();
        _output.WriteLine($"Baseline POWERPNT.exe PIDs before test: [{string.Join(", ", baselinePowerPointPids)}]");

        var tempDir = Path.Join(Path.GetTempPath(), $"McpShutdownRobustnessTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        _output.WriteLine($"Test directory: {tempDir}");

        var clientToServerPipe = new Pipe();
        var serverToClientPipe = new Pipe();
        using var cts = new CancellationTokenSource();

        McpClient? client = null;
        Task? serverTask = null;

        try
        {
            (client, serverTask) = await ProgramTransportTestHost.StartAsync(
                clientToServerPipe,
                serverToClientPipe,
                "ShutdownRobustnessTestClient",
                cts.Token);
            _output.WriteLine($"✓ Connected to server: {client.ServerInfo?.Name} v{client.ServerInfo?.Version}");

            Task<string> Call(string toolName, Dictionary<string, object?> arguments)
                => CallToolAsync(client, toolName, arguments, cts.Token);

            // Pay the one, unavoidable create_presentation cost once; reuse the resulting file via
            // plain filesystem copies (no COM) for every session this test needs. create_presentation
            // now creates-and-keeps-open (returns a live sessionId), so we must close that seed
            // session once the file is on disk — otherwise it lingers in the registry and the
            // final list_sessions count assertion (expecting 0) would see it.
            var seedFile = Path.Join(tempDir, "Seed.pptx");
            var seedCreateJson = await Call("create_presentation", new() { ["filePath"] = seedFile });
            AssertSuccess(seedCreateJson, "create_presentation");
            var seedSessionId = GetString(seedCreateJson, "sessionId");
            Assert.True(File.Exists(seedFile));
            _output.WriteLine("✓ create_presentation (seed file)");

            string fileA = Path.Join(tempDir, "A.pptx");
            string fileB = Path.Join(tempDir, "B.pptx");
            string fileC = Path.Join(tempDir, "C.pptx");
            string fileD = Path.Join(tempDir, "D.pptx");
            foreach (var f in new[] { fileA, fileB, fileC, fileD })
            {
                File.Copy(seedFile, f);
            }

            // Close the seed session now that its file has been copied — the test only needs the
            // file on disk, not a live seed session. Its background disposal is tracked and awaited
            // by DisposeAll on host shutdown like any other session.
            AssertSuccess(await Call("close_presentation", new() { ["sessionId"] = seedSessionId }), "close_presentation (seed)");
            _output.WriteLine("✓ close_presentation (seed session)");

            // Open sessions sequentially — concurrent PowerPoint COM activation is transiently
            // flaky (Ripley's charter: "RPC server is unavailable" under concurrent launch load).
            var sessionA = await OpenSession(Call, fileA, "A");
            var sessionB = await OpenSession(Call, fileB, "B");
            var sessionC = await OpenSession(Call, fileC, "C");
            var sessionD = await OpenSession(Call, fileD, "D");

            // --- Double-close: second close of the same session is graceful, not an exception. ---
            var firstClose = await TimedClose(Call, sessionA);
            AssertSuccess(firstClose.Json, "close_presentation A (1st)");
            Assert.True(GetBool(firstClose.Json, "closed"));
            Assert.True(firstClose.Elapsed < TimeSpan.FromSeconds(15), $"1st close of A should be fast; took {firstClose.Elapsed}.");

            var secondClose = await TimedClose(Call, sessionA);
            using (var secondCloseJson = JsonDocument.Parse(secondClose.Json))
            {
                Assert.True(
                    secondCloseJson.RootElement.TryGetProperty("success", out var successProp) && !successProp.GetBoolean(),
                    $"Second close of an already-closed session should return success=false, not throw or hang: {secondClose.Json}");
                Assert.True(
                    secondCloseJson.RootElement.TryGetProperty("errorMessage", out var errProp)
                        && (errProp.GetString()?.Contains("Unknown sessionId", StringComparison.OrdinalIgnoreCase) ?? false),
                    $"Expected a graceful 'unknown sessionId' error on double-close: {secondClose.Json}");
            }
            Assert.True(secondClose.Elapsed < TimeSpan.FromSeconds(15), $"2nd close of A should be fast, not hang; took {secondClose.Elapsed}.");
            _output.WriteLine($"✓ Double-close: 1st close={firstClose.Elapsed.TotalMilliseconds:N0}ms (closed=true), 2nd close={secondClose.Elapsed.TotalMilliseconds:N0}ms (graceful not-found)");

            // --- Concurrent close: closing two independent sessions at the same time. ---
            var concurrentStopwatch = Stopwatch.StartNew();
            var closeBTask = Call("close_presentation", new() { ["sessionId"] = sessionB });
            var closeCTask = Call("close_presentation", new() { ["sessionId"] = sessionC });
            var closeResults = await Task.WhenAll(closeBTask, closeCTask);
            concurrentStopwatch.Stop();

            AssertSuccess(closeResults[0], "close_presentation B (concurrent)");
            AssertSuccess(closeResults[1], "close_presentation C (concurrent)");
            Assert.True(GetBool(closeResults[0], "closed"));
            Assert.True(GetBool(closeResults[1], "closed"));
            Assert.True(
                concurrentStopwatch.Elapsed < TimeSpan.FromSeconds(15),
                $"Concurrent closes of B and C should both return fast; took {concurrentStopwatch.Elapsed}.");
            _output.WriteLine($"✓ Concurrent close of B and C completed in {concurrentStopwatch.ElapsedMilliseconds}ms (both closed=true)");

            // --- Close-then-immediate-host-shutdown race: no artificial delay between the async
            // close_presentation call returning and the host being asked to stop. DisposeAll must
            // still pick up D's (and A/B/C's) in-flight background disposal before the process
            // fully exits. ---
            var closeD = await TimedClose(Call, sessionD);
            AssertSuccess(closeD.Json, "close_presentation D");
            Assert.True(GetBool(closeD.Json, "closed"));
            Assert.True(closeD.Elapsed < TimeSpan.FromSeconds(15), $"Close of D should be fast; took {closeD.Elapsed}.");
            _output.WriteLine($"✓ close_presentation D returned in {closeD.Elapsed.TotalMilliseconds:N0}ms; stopping host immediately (no delay)");

            var listAfterAllClosesResult = await Call("list_sessions", []);
            AssertSuccess(listAfterAllClosesResult, "list_sessions (after all closes)");
            using (var listJson = JsonDocument.Parse(listAfterAllClosesResult))
            {
                Assert.Equal(0, listJson.RootElement.GetProperty("count").GetInt32());
            }
            _output.WriteLine("✓ list_sessions confirms all four sessions are gone immediately after their (async) closes");
        }
        finally
        {
            // Stopping the host drives PresentationSessionShutdownService.StopAsync →
            // registry.DisposeAll(), which blocks until every in-flight background disposal
            // (A, B, C, D — regardless of how recently each was closed) has actually finished.
            var shutdownStopwatch = Stopwatch.StartNew();
            await ProgramTransportTestHost.StopAsync(client, clientToServerPipe, serverToClientPipe, serverTask, _output, cts);
            shutdownStopwatch.Stop();
            _output.WriteLine($"✓ Host shutdown (awaiting all pending disposals) took {shutdownStopwatch.Elapsed}");

            // The critical assertion: after the host has FULLY stopped, no POWERPNT.exe process
            // that wasn't already running before this test (baseline) may still be alive. Diffing
            // against the baseline (rather than asserting zero total) tolerates unrelated
            // processes from other sources while still catching anything THIS test leaked.
            var survivingPids = GetCurrentPowerPointPids().Except(baselinePowerPointPids).ToList();
            Assert.True(
                survivingPids.Count == 0,
                $"Orphaned POWERPNT.exe process(es) survived host shutdown: [{string.Join(", ", survivingPids)}]");
            _output.WriteLine("✓ No orphaned POWERPNT.exe processes from this test remain after host shutdown");

            if (Directory.Exists(tempDir))
            {
                try
                {
                    Directory.Delete(tempDir, recursive: true);
                }
                catch
                {
                    // Best-effort cleanup only.
                }
            }
        }
    }

    private static async Task<string> OpenSession(
        Func<string, Dictionary<string, object?>, Task<string>> call,
        string filePath,
        string label)
    {
        var result = await call("open_presentation", new Dictionary<string, object?> { ["filePath"] = filePath });
        AssertSuccess(result, $"open_presentation ({label})");
        var sessionId = GetString(result, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId), $"Expected a sessionId for session {label}: {result}");
        return sessionId!;
    }

    private static async Task<(string Json, TimeSpan Elapsed)> TimedClose(
        Func<string, Dictionary<string, object?>, Task<string>> call,
        string sessionId)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = await call("close_presentation", new Dictionary<string, object?> { ["sessionId"] = sessionId });
        stopwatch.Stop();
        return (result, stopwatch.Elapsed);
    }

    private static List<int> GetCurrentPowerPointPids()
    {
        var result = new List<int>();
        foreach (var process in Process.GetProcessesByName("POWERPNT"))
        {
            using (process)
            {
                result.Add(process.Id);
            }
        }

        return result;
    }
}
