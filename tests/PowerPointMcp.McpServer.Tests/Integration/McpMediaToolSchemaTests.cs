using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using Sbroenne.PowerPointMcp.Generated;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Verifies generated MCP schema and CLI routing for <c>IMediaCommands</c>. These tests exercise
/// generated contracts only; real media COM behavior is covered by Core integration tests.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Media")]
public sealed class McpMediaToolSchemaTests : IAsyncLifetime, IAsyncDisposable
{
    private static readonly HashSet<string> ExpectedParameters = new(StringComparer.Ordinal)
    {
        "action",
        "session_id",
        "slide_index",
        "media_path",
        "link_to_file",
        "save_with_document",
        "left",
        "top",
        "width",
        "height",
        "shape_index",
    };

    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpMediaToolSchemaTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "MediaSchemaTestClient",
            _cts.Token);
    }

    public async Task DisposeAsync() => await DisposeAsyncCore();

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void MediaAction_EnumAndCliRouting_AreComplete()
    {
        string[] expectedActions = ["add-media", "get-media-info"];
        Assert.Equal(expectedActions, ServiceRegistry.Media.ValidActions);

        var enumActions = Enum.GetValues<MediaAction>()
            .Select(ServiceRegistry.Media.ToActionString)
            .ToArray();
        Assert.Equal(expectedActions, enumActions);

        var (addCommand, _) = ServiceRegistry.Media.RouteCliArgs(
            "add-media",
            slideIndex: 1,
            mediaPath: "C:\\media.wav",
            linkToFile: false,
            saveWithDocument: true,
            left: 0,
            top: 0,
            width: 100,
            height: 100);
        Assert.Equal("media.add-media", addCommand);

        var (getCommand, _) = ServiceRegistry.Media.RouteCliArgs(
            "get-media-info",
            slideIndex: 1,
            shapeIndex: 1);
        Assert.Equal("media.get-media-info", getCommand);
    }

    [Fact]
    public async Task MediaTool_SchemaContainsGeneratedActionsAndParameters()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var mediaTool = tools.Single(tool => tool.Name == ServiceRegistry.Media.McpToolName);

        Assert.Contains("add-media", mediaTool.Description, StringComparison.Ordinal);
        Assert.Contains("get-media-info", mediaTool.Description, StringComparison.Ordinal);

        JsonElement schema = mediaTool.JsonSchema;
        JsonElement properties = schema.GetProperty("properties");
        var actualParameters = properties.EnumerateObject()
            .Select(property => property.Name)
            .ToHashSet(StringComparer.Ordinal);
        Assert.True(
            ExpectedParameters.SetEquals(actualParameters),
            $"Expected [{string.Join(", ", ExpectedParameters.Order())}], got " +
            $"[{string.Join(", ", actualParameters.Order())}].");

        JsonElement actionSchema = properties.GetProperty("action");
        string[] actionValues = actionSchema.GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString()!)
            .ToArray();
        Assert.Equal(ServiceRegistry.Media.ValidActions, actionValues);
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
}
