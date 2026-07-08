// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using ModelContextProtocol.Client;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Protocol-level proof that the PowerPoint MCP server is wired correctly — no PowerPoint
/// required. Drives the real <see cref="Program"/> host over an in-memory pipe transport and
/// asserts the tool surface via the official MCP SDK client (tools/list + schema shape), never
/// via reflection or direct method calls.
/// </summary>
/// <remarks>
/// This test stands alone: it must pass in any environment, including one without PowerPoint
/// installed, because it never calls a tool that touches COM. Live-COM round-trip coverage lives
/// in <see cref="McpRoundTripTests"/>.
/// </remarks>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public sealed class McpProtocolTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    /// <summary>
    /// The full 33-tool MCP surface across all 10 domains (Presentation, Slide, Shape, TextFrame,
    /// Table, Notes, Layout, Image, Chart, Export) — the source of truth for this test, enumerated
    /// directly from every <c>[McpServerTool]</c> in <c>src/PowerPointMcp.McpServer/Tools/*.cs</c>.
    /// If this set changes, update it deliberately alongside the tool surface (see
    /// .squad/decisions/inbox/brett-remaining-domain-tools.md for the 5→31 tool-count change, and
    /// .squad/decisions/inbox/copilot-issue3-moduleinitializer-root-cause.md for the 31→33 change
    /// adding apply_template/get_theme_name for .potx template-apply support).
    /// </summary>
    private static readonly HashSet<string> ExpectedToolNames =
    [
        // PresentationTools.cs (7)
        "create_presentation",
        "open_presentation",
        "save_presentation",
        "close_presentation",
        "list_sessions",
        "apply_template",
        "get_theme_name",
        // SlideTools.cs (3)
        "add_slide",
        "get_slide_count",
        "delete_slide",
        // ShapeTools.cs (6)
        "add_rectangle",
        "add_text_box",
        "get_shape_count",
        "delete_shape",
        "set_shape_position",
        "set_shape_size",
        // TextFrameTools.cs (5)
        "set_text",
        "get_text",
        "set_font_size",
        "set_bold",
        "set_font_color",
        // TableTools.cs (3)
        "add_table",
        "set_cell_text",
        "get_cell_text",
        // NotesTools.cs (2)
        "set_notes_text",
        "get_notes_text",
        // LayoutTools.cs (2)
        "set_layout",
        "get_layout",
        // ImageTools.cs (1)
        "add_picture",
        // ChartTools.cs (2)
        "add_chart",
        "get_chart_data",
        // ExportTools.cs (2)
        "export_slide_to_image",
        "export_all_slides_to_images"
    ];

    public McpProtocolTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "ProtocolTestClient",
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
    }

    /// <summary>
    /// THE core protocol proof: exactly the 33 expected tools (across all 10 domains) are
    /// discoverable via <c>tools/list</c> — no more, no less.
    /// </summary>
    [Fact]
    public async Task ListTools_ReturnsExactlyTheThirtyThreeExpectedTools()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        _output.WriteLine($"Discovered {tools.Count} tools via MCP protocol:");
        foreach (var tool in tools.OrderBy(t => t.Name, StringComparer.Ordinal))
        {
            _output.WriteLine($"  • {tool.Name}: {tool.Description}");
        }

        var actualToolNames = tools.Select(t => t.Name).ToHashSet();

        var missingTools = ExpectedToolNames.Except(actualToolNames).ToList();
        Assert.True(missingTools.Count == 0, $"Missing tools: {string.Join(", ", missingTools)}");

        var unexpectedTools = actualToolNames.Except(ExpectedToolNames).ToList();
        Assert.True(unexpectedTools.Count == 0, $"Unexpected tools: {string.Join(", ", unexpectedTools)}");

        Assert.Equal(ExpectedToolNames.Count, tools.Count);
    }

    /// <summary>
    /// The DI-injected <c>PresentationSessionRegistry registry</c> parameter on
    /// open_presentation/save_presentation/close_presentation/list_sessions must never leak into
    /// the JSON schema the client sees — it's satisfied from the host's service provider, not
    /// supplied by the caller.
    /// </summary>
    [Fact]
    public async Task ListTools_NoToolSchemaExposesTheRegistryParameter()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        Assert.NotEmpty(tools);

        foreach (var tool in tools)
        {
            var schema = tool.JsonSchema;
            if (!schema.TryGetProperty("properties", out var properties))
            {
                continue;
            }

            foreach (var property in properties.EnumerateObject())
            {
                Assert.False(
                    string.Equals(property.Name, "registry", StringComparison.OrdinalIgnoreCase),
                    $"Tool '{tool.Name}' leaked the DI-injected 'registry' parameter into its schema: {schema.GetRawText()}");
            }

            _output.WriteLine($"✓ {tool.Name}: schema has no 'registry' parameter ({properties.EnumerateObject().Count()} properties)");
        }
    }

    /// <summary>
    /// Sanity check that every tool has a name and description — cheap and catches attribute
    /// mistakes ([McpServerTool(Name=...)] typos, missing [Description]) early.
    /// </summary>
    [Fact]
    public async Task ListTools_AllToolsHaveNameAndDescription()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        foreach (var tool in tools)
        {
            Assert.False(string.IsNullOrEmpty(tool.Name), "Tool has empty name");
            Assert.False(string.IsNullOrEmpty(tool.Description), $"Tool {tool.Name} has no description");
        }
    }

    /// <summary>
    /// Server info/instructions surfaced via the MCP protocol match Program.cs's configuration.
    /// </summary>
    [Fact]
    public void ServerInfo_ReturnsCorrectInformation()
    {
        var serverInfo = _client!.ServerInfo;

        Assert.NotNull(serverInfo);
        Assert.Equal("powerpoint-mcp", serverInfo.Name);
        Assert.NotNull(_client.ServerInstructions);
        Assert.Contains("create_presentation", _client.ServerInstructions, StringComparison.Ordinal);
    }
}
