// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// End-to-end round-trip test that drives the real MCP host over an in-memory pipe transport and
/// exercises the full Presentation session lifecycle against a REAL PowerPoint instance (Rule 30:
/// real COM, no mocking). The server is treated as a black box — every assertion goes through the
/// MCP protocol (tools/list, tools/call), never through direct method calls.
/// </summary>
/// <remarks>
/// This is separate from <see cref="McpProtocolTests"/> so that the schema/protocol proof stands
/// alone even in an environment without PowerPoint installed. This class DOES require PowerPoint;
/// if PowerPoint COM activation fails, this test fails loudly (no silent mock fallback) — that is
/// the point of Rule 30. Runs in the "ProgramTransport" collection so it never overlaps with
/// another transport-hook test or another real PowerPoint COM launch.
/// </remarks>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpRoundTripTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testPresentationFile;

    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpRoundTripTests(ITestOutputHelper output)
    {
        _output = output;

        _tempDir = Path.Join(Path.GetTempPath(), $"McpRoundTripTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testPresentationFile = Path.Join(_tempDir, "RoundTrip.pptx");

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "RoundTripTestClient",
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
    /// create_presentation via the MCP protocol: real PowerPoint creates and saves a real .pptx
    /// on disk and returns an OPEN session (create-and-keep-open semantics). The session is closed
    /// at the end so no PowerPoint instance is left running.
    /// </summary>
    [Fact]
    public async Task CreatePresentation_ViaMcpProtocol_WritesRealPptxFile()
    {
        var result = await CallToolAsync("presentation", new Dictionary<string, object?>
        { ["action"] = "create",
            ["filePath"] = _testPresentationFile
        });

        AssertSuccess(result, "create_presentation");
        Assert.True(File.Exists(_testPresentationFile), $"Expected file to exist: {_testPresentationFile}");
        Assert.Equal(_testPresentationFile, GetJsonProperty(result, "presentationPath"));

        var sessionId = GetJsonProperty(result, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId), $"Expected create_presentation to return an open sessionId: {result}");

        // create-and-keep-open: close the returned session so no PowerPoint instance leaks.
        var closeResult = await CallToolAsync("presentation", new Dictionary<string, object?>
        { ["action"] = "close",
            ["sessionId"] = sessionId
        });
        AssertSuccess(closeResult, "close_presentation");

        _output.WriteLine($"✓ create_presentation wrote a real .pptx file and returned sessionId={sessionId}: {_testPresentationFile}");
    }

    /// <summary>
    /// Full session lifecycle via the MCP protocol: create → open → list_sessions shows it →
    /// save → close. Every step is asserted through tools/call responses only.
    /// </summary>
    [Fact]
    public async Task FullSessionLifecycle_ViaMcpProtocol_OpenListSaveClose()
    {
        // 1. create_presentation returns an OPEN session (create-and-keep-open) → sessionId.
        var createResult = await CallToolAsync("presentation", new Dictionary<string, object?>
        { ["action"] = "create",
            ["filePath"] = _testPresentationFile
        });
        AssertSuccess(createResult, "create_presentation");
        Assert.True(File.Exists(_testPresentationFile));
        var sessionId = GetJsonProperty(createResult, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId), $"Expected a sessionId in create response: {createResult}");
        _output.WriteLine($"✓ Step 1: create_presentation returned open sessionId={sessionId}");

        // 2. list_sessions shows the open session.
        var listResult = await CallToolAsync("presentation", new() { ["action"] = "list" });
        AssertSuccess(listResult, "list_sessions");
        using (var listJson = JsonDocument.Parse(listResult))
        {
            var sessions = listJson.RootElement.GetProperty("sessions");
            var found = sessions.EnumerateArray()
                .Any(s => string.Equals(s.GetProperty("sessionId").GetString(), sessionId, StringComparison.Ordinal));
            Assert.True(found, $"Expected sessionId {sessionId} in list_sessions response: {listResult}");
        }
        _output.WriteLine("✓ Step 2: list_sessions shows the open session");

        // 3. save_presentation.
        var saveResult = await CallToolAsync("presentation", new Dictionary<string, object?>
        { ["action"] = "save",
            ["sessionId"] = sessionId
        });
        AssertSuccess(saveResult, "save_presentation");
        _output.WriteLine("✓ Step 3: save_presentation succeeded");

        // 4. close_presentation — releases the PowerPoint process for this session.
        var closeResult = await CallToolAsync("presentation", new Dictionary<string, object?>
        { ["action"] = "close",
            ["sessionId"] = sessionId
        });
        AssertSuccess(closeResult, "close_presentation");
        using (var closeJson = JsonDocument.Parse(closeResult))
        {
            Assert.True(closeJson.RootElement.GetProperty("closed").GetBoolean(), $"Expected closed=true: {closeResult}");
        }
        _output.WriteLine("✓ Step 4: close_presentation succeeded");

        // 5. list_sessions no longer shows the closed session.
        var listAfterCloseResult = await CallToolAsync("presentation", new() { ["action"] = "list" });
        AssertSuccess(listAfterCloseResult, "list_sessions (after close)");
        using (var listJson = JsonDocument.Parse(listAfterCloseResult))
        {
            var sessions = listJson.RootElement.GetProperty("sessions");
            var stillFound = sessions.EnumerateArray()
                .Any(s => string.Equals(s.GetProperty("sessionId").GetString(), sessionId, StringComparison.Ordinal));
            Assert.False(stillFound, $"Session {sessionId} should be gone after close_presentation: {listAfterCloseResult}");
        }
        _output.WriteLine("✓ Step 5: list_sessions confirms the session is closed");
    }

    private async Task<string> CallToolAsync(string toolName, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(toolName, arguments, cancellationToken: _cts.Token);

        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        return textBlock.Text;
    }

    private static void AssertSuccess(string jsonResult, string operationName)
    {
        JsonDocument json;
        try
        {
            json = JsonDocument.Parse(jsonResult);
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
            return;
        }

        using (json)
        {
            if (json.RootElement.TryGetProperty("success", out var success) && !success.GetBoolean())
            {
                var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                    ? errProp.GetString()
                    : "Unknown error";
                Assert.Fail($"{operationName} returned success=false: {errorMsg}\nResponse: {jsonResult}");
            }
        }
    }

    private static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }
}
