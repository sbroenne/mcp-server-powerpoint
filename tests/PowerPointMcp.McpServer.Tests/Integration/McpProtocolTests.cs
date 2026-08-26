// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
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
    /// The MCP tool surface: one hand-written action-dispatch tool (Presentation — session
    /// lifecycle + template + document properties) plus one generated action-dispatch tool per
    /// remaining Core domain (Slide, Shape, TextFrame, Table, Notes, Layout, PageSetup,
    /// Accessibility, Master, Animation, SmartArt, Image, Chart, Export) — enumerated directly
    /// from every <c>[McpServerTool]</c> in
    /// <c>src/PowerPointMcp.McpServer/Tools/*.cs</c> (hand-written) and the generated
    /// <c>PowerPointMcp.Generators.Mcp</c> output (one action-dispatch tool per domain, matching
    /// mcp-server-excel's architecture: a single tool per domain with an action enum, instead of
    /// one tool per verb). If this set changes, update it deliberately alongside the tool surface.
    /// </summary>
    private static readonly HashSet<string> ExpectedToolNames =
    [
        // PresentationTools.cs (1, hand-written action-dispatch tool — session lifecycle,
        // Save As/copy, template, Mark as Final, and document properties; 16 actions)
        "presentation",
        // Generated action-dispatch tools (14, one per remaining Core domain)
        "slide",
        "shape",
        "textframe",
        "table",
        "notes",
        "layout",
        "pagesetup",
        "accessibility",
        "master",
        "animation",
        "smartart",
        "image",
        "chart",
        "export"
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
    /// THE core protocol proof: exactly the 15 expected tools (1 hand-written + 14 generated
    /// action-dispatch tools) are discoverable via <c>tools/list</c> — no more, no less.
    /// </summary>
    [Fact]
    public async Task ListTools_ReturnsExactlyTheExpectedTools()
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
    /// The DI-injected <c>PresentationSessionRegistry registry</c> parameter (hand-written
    /// session-lifecycle tools) and <c>PowerPointMcpService service</c> parameter (generated
    /// action-dispatch tools) must never leak into the JSON schema the client sees — both are
    /// satisfied from the host's service provider, not supplied by the caller.
    /// </summary>
    [Fact]
    public async Task ListTools_NoToolSchemaExposesDiInjectedParameters()
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
                    string.Equals(property.Name, "registry", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property.Name, "service", StringComparison.OrdinalIgnoreCase),
                    $"Tool '{tool.Name}' leaked a DI-injected parameter into its schema: {schema.GetRawText()}");
            }

            _output.WriteLine($"✓ {tool.Name}: schema has no DI-injected parameter ({properties.EnumerateObject().Count()} properties)");
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

    [Fact]
    public async Task PresentationSchema_UsesCanonicalLifecycleActions()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var presentation = Assert.Single(tools, tool => tool.Name == "presentation");
        var actionSchema = presentation.JsonSchema
            .GetProperty("properties")
            .GetProperty("action");
        var actions = actionSchema
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("test", actions);
        Assert.Contains("save-as", actions);
        Assert.Contains("save-copy-as", actions);
        Assert.DoesNotContain("save", actions);

        var properties = presentation.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("save", out var saveSchema));
        var saveTypes = saveSchema.GetProperty("type")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);
        Assert.Contains("boolean", saveTypes);
        Assert.Contains("null", saveTypes);

        Assert.True(properties.TryGetProperty("targetPath", out _));
        Assert.True(properties.TryGetProperty("overwrite", out _));
        var formatValues = properties
            .GetProperty("format")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .Where(value => value != null)
            .ToHashSet(StringComparer.Ordinal);
        Assert.Equal(
            new HashSet<string?> { "auto", "pptx", "pptm", "ppt" },
            formatValues);
    }

    [Fact]
    public async Task PresentationSchema_ExposesFinalActionsAndAdvisoryContract()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var presentation = Assert.Single(tools, tool => tool.Name == "presentation");
        var properties = presentation.JsonSchema.GetProperty("properties");
        var actions = properties
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("get-final", actions);
        Assert.Contains("set-final", actions);

        var isFinalSchema = properties.GetProperty("isFinal");
        var isFinalTypes = isFinalSchema
            .GetProperty("type")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);
        Assert.Contains("boolean", isFinalTypes);
        Assert.Contains("null", isFinalTypes);

        var contractText = $"{presentation.Description} {isFinalSchema.GetProperty("description").GetString()}";
        Assert.Contains("advisory", contractText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not authentication", contractText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("encryption", contractText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("access control", contractText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PresentationFinalActions_RejectUnknownSessionsAndInapplicableParameters()
    {
        var unknownGet = await CallPresentationAsync(new()
        {
            ["action"] = "get-final",
            ["sessionId"] = "missing-session"
        });
        Assert.False(unknownGet.GetProperty("success").GetBoolean());
        Assert.Contains("Unknown sessionId", unknownGet.GetProperty("errorMessage").GetString());

        var unknownSet = await CallPresentationAsync(new()
        {
            ["action"] = "set-final",
            ["sessionId"] = "missing-session",
            ["isFinal"] = true
        });
        Assert.False(unknownSet.GetProperty("success").GetBoolean());
        Assert.Contains("Unknown sessionId", unknownSet.GetProperty("errorMessage").GetString());

        var missingValue = await CallPresentationAsync(new()
        {
            ["action"] = "set-final",
            ["sessionId"] = "missing-session"
        });
        Assert.False(missingValue.GetProperty("success").GetBoolean());
        Assert.Contains("isFinal is required", missingValue.GetProperty("errorMessage").GetString());

        var inapplicableValue = await CallPresentationAsync(new()
        {
            ["action"] = "get-final",
            ["sessionId"] = "missing-session",
            ["isFinal"] = false
        });
        Assert.False(inapplicableValue.GetProperty("success").GetBoolean());
        Assert.Contains("not valid for action 'get-final'", inapplicableValue.GetProperty("errorMessage").GetString());
    }

    private async Task<JsonElement> CallPresentationAsync(Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync("presentation", arguments, cancellationToken: _cts.Token);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var json = JsonDocument.Parse(text);
        return json.RootElement.Clone();
    }

    [Fact]
    public async Task TagSchemas_ExposeAllOwnerActionsAndStrictParameters()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        var presentation = Assert.Single(tools, tool => tool.Name == "presentation");
        AssertTagSchema(
            presentation.JsonSchema,
            expectedProperties: ["action", "sessionId", "tagName", "tagValue"],
            forbiddenProperties: ["slide_index", "shape_index", "binary_value"]);

        var slide = Assert.Single(tools, tool => tool.Name == "slide");
        AssertTagSchema(
            slide.JsonSchema,
            expectedProperties: ["action", "session_id", "slide_index", "tag_name", "tag_value"],
            forbiddenProperties: ["shape_index", "binary_value"]);

        var shape = Assert.Single(tools, tool => tool.Name == "shape");
        AssertTagSchema(
            shape.JsonSchema,
            expectedProperties: ["action", "session_id", "slide_index", "shape_index", "tag_name", "tag_value"],
            forbiddenProperties: ["binary_value"]);
    }

    private static void AssertTagSchema(
        System.Text.Json.JsonElement schema,
        IReadOnlyList<string> expectedProperties,
        IReadOnlyList<string> forbiddenProperties)
    {
        var actions = schema
            .GetProperty("properties")
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("set-tag", actions);
        Assert.Contains("get-tag", actions);
        Assert.Contains("list-tags", actions);
        Assert.Contains("delete-tag", actions);

        var properties = schema.GetProperty("properties");
        foreach (string property in expectedProperties)
        {
            Assert.True(properties.TryGetProperty(property, out _), $"Expected schema property '{property}'.");
        }

        foreach (string property in forbiddenProperties)
        {
            Assert.False(properties.TryGetProperty(property, out _), $"Unexpected schema property '{property}'.");
        }
    }

    [Fact]
    public async Task ChartSchema_ExposesQuickFormattingActionsAndParameters()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var chart = Assert.Single(tools, tool => tool.Name == "chart");
        var properties = chart.JsonSchema.GetProperty("properties");
        var actions = properties
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("get-style", actions);
        Assert.Contains("set-style", actions);
        Assert.Contains("get-color-style", actions);
        Assert.Contains("set-color-style", actions);
        Assert.Contains("get-data-table", actions);
        Assert.Contains("set-data-table", actions);
        Assert.True(properties.TryGetProperty("style", out _));
        Assert.True(properties.TryGetProperty("color_style", out _));
        var visible = properties.GetProperty("visible");
        string visibleDescription = visible.GetProperty("description").GetString()!;
        Assert.Contains("set-legend-visibility", visibleDescription);
        Assert.Contains("set-data-table", visibleDescription);
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
        Assert.Contains("presentation(action=create", _client.ServerInstructions, StringComparison.Ordinal);
    }
}
