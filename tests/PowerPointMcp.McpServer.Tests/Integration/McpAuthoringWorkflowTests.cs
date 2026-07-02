// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using ModelContextProtocol.Client;
using Xunit.Abstractions;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Integration.McpToolCallHelper;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// One realistic, end-to-end deck-authoring workflow driven ENTIRELY through the MCP protocol
/// (tools/call over the in-memory client) — a single session that touches every domain tool once,
/// to amortize the real cost of a live PowerPoint COM session (Rule 30: no mocking).
/// </summary>
/// <remarks>
/// Deliberately a single long test rather than many small ones: <c>create_presentation</c> alone
/// currently blocks synchronously for the full PowerPoint quit/grace-period sequence (observed
/// ~100-200s — see .squad/decisions/inbox/ripley-create-presentation-blocks-on-dispose.md), so
/// paying that cost more than once per test class would make the suite unreasonably slow. Every
/// assertion here goes through a tools/call JSON response — never a direct method call — per
/// Ripley's charter (integration tests only, no mocking, real COM).
/// </remarks>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "DeckAuthoring")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpAuthoringWorkflowTests : IAsyncLifetime, IAsyncDisposable
{
    // Minimal valid 1x1 black pixel PNG, used to prove add_picture round-trips through real COM
    // without needing System.Drawing or any extra package reference.
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    private static readonly string[] ChartCategories = ["Q1", "Q2", "Q3"];
    private static readonly double[] ChartValues = [10.0, 20.0, 30.0];

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _presentationFile;
    private readonly string _imageFile;

    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpAuthoringWorkflowTests(ITestOutputHelper output)
    {
        _output = output;

        _tempDir = Path.Join(Path.GetTempPath(), $"McpAuthoringWorkflowTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _presentationFile = Path.Join(_tempDir, "Authoring.pptx");
        _imageFile = Path.Join(_tempDir, "pixel.png");

        File.WriteAllBytes(_imageFile, Convert.FromBase64String(OnePixelPngBase64));

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "AuthoringWorkflowTestClient",
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
                // Best-effort cleanup only.
            }
        }
    }

    /// <summary>
    /// Full realistic deck-authoring workflow in one session, exercising every domain's MCP tools
    /// through tools/call and asserting only via the returned JSON.
    /// </summary>
    [Fact]
    public async Task FullDeckAuthoringWorkflow_ViaMcpProtocol_ExercisesEveryDomainTool()
    {
        // 1. create_presentation → open_presentation → sessionId.
        var createResult = await Call("create_presentation", new() { ["filePath"] = _presentationFile });
        AssertSuccess(createResult, "create_presentation");
        Assert.True(File.Exists(_presentationFile));
        _output.WriteLine("✓ create_presentation");

        var openResult = await Call("open_presentation", new() { ["filePath"] = _presentationFile });
        AssertSuccess(openResult, "open_presentation");
        var sessionId = GetString(openResult, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId));
        _output.WriteLine($"✓ open_presentation → sessionId={sessionId}");

        // 2. add_slide (x2), get_slide_count asserts the count grew by exactly 2.
        var baselineCountResult = await Call("get_slide_count", new() { ["sessionId"] = sessionId });
        AssertSuccess(baselineCountResult, "get_slide_count (baseline)");
        var baselineCount = GetInt(baselineCountResult, "slideCount")!.Value;

        AssertSuccess(await Call("add_slide", new() { ["sessionId"] = sessionId }), "add_slide #1");
        AssertSuccess(await Call("add_slide", new() { ["sessionId"] = sessionId }), "add_slide #2");

        var afterAddCountResult = await Call("get_slide_count", new() { ["sessionId"] = sessionId });
        AssertSuccess(afterAddCountResult, "get_slide_count (after add)");
        var afterAddCount = GetInt(afterAddCountResult, "slideCount")!.Value;
        Assert.Equal(baselineCount + 2, afterAddCount);
        _output.WriteLine($"✓ add_slide x2, get_slide_count {baselineCount} → {afterAddCount}");

        const int slideIndex = 1;

        // 3. add_text_box / set_text / set_font_size / set_bold / set_font_color; get_text round-trips.
        var addTextBoxResult = await Call("add_text_box", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["left"] = 10f,
            ["top"] = 10f,
            ["width"] = 200f,
            ["height"] = 50f,
            ["text"] = "Hello"
        });
        AssertSuccess(addTextBoxResult, "add_text_box");
        var textBoxShapeIndex = GetInt(addTextBoxResult, "shapeIndex")!.Value;
        _output.WriteLine($"✓ add_text_box → shapeIndex={textBoxShapeIndex}");

        AssertSuccess(await Call("set_text", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = textBoxShapeIndex,
            ["text"] = "Updated Text"
        }), "set_text");

        AssertSuccess(await Call("set_font_size", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = textBoxShapeIndex,
            ["fontSize"] = 24f
        }), "set_font_size");

        AssertSuccess(await Call("set_bold", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = textBoxShapeIndex,
            ["bold"] = true
        }), "set_bold");

        AssertSuccess(await Call("set_font_color", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = textBoxShapeIndex,
            ["red"] = (byte)255,
            ["green"] = (byte)0,
            ["blue"] = (byte)0
        }), "set_font_color");

        var getTextResult = await Call("get_text", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = textBoxShapeIndex
        });
        AssertSuccess(getTextResult, "get_text");
        Assert.Equal("Updated Text", GetString(getTextResult, "text"));
        _output.WriteLine("✓ set_text/set_font_size/set_bold/set_font_color, get_text round-trips");

        // 4. add_rectangle, set_shape_position, set_shape_size, get_shape_count, delete a shape.
        var addRectResult = await Call("add_rectangle", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["left"] = 50f,
            ["top"] = 80f,
            ["width"] = 100f,
            ["height"] = 60f
        });
        AssertSuccess(addRectResult, "add_rectangle");
        var rectShapeIndex = GetInt(addRectResult, "shapeIndex")!.Value;

        AssertSuccess(await Call("set_shape_position", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = rectShapeIndex,
            ["left"] = 75f,
            ["top"] = 90f
        }), "set_shape_position");

        AssertSuccess(await Call("set_shape_size", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = rectShapeIndex,
            ["width"] = 120f,
            ["height"] = 70f
        }), "set_shape_size");

        var shapeCountBeforeDeleteResult = await Call("get_shape_count", new() { ["sessionId"] = sessionId, ["slideIndex"] = slideIndex });
        AssertSuccess(shapeCountBeforeDeleteResult, "get_shape_count (before delete)");
        var shapeCountBeforeDelete = GetInt(shapeCountBeforeDeleteResult, "shapeCount")!.Value;

        AssertSuccess(await Call("delete_shape", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = rectShapeIndex
        }), "delete_shape");

        var shapeCountAfterDeleteResult = await Call("get_shape_count", new() { ["sessionId"] = sessionId, ["slideIndex"] = slideIndex });
        AssertSuccess(shapeCountAfterDeleteResult, "get_shape_count (after delete)");
        var shapeCountAfterDelete = GetInt(shapeCountAfterDeleteResult, "shapeCount")!.Value;
        Assert.Equal(shapeCountBeforeDelete - 1, shapeCountAfterDelete);
        _output.WriteLine($"✓ add_rectangle/set_shape_position/set_shape_size/delete_shape, shapeCount {shapeCountBeforeDelete} → {shapeCountAfterDelete}");

        // 5. add_table, set_cell_text, get_cell_text round-trip.
        var addTableResult = await Call("add_table", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["rows"] = 2,
            ["columns"] = 2,
            ["left"] = 20f,
            ["top"] = 200f,
            ["width"] = 300f,
            ["height"] = 100f
        });
        AssertSuccess(addTableResult, "add_table");
        var tableShapeIndex = GetInt(addTableResult, "shapeIndex")!.Value;

        AssertSuccess(await Call("set_cell_text", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = tableShapeIndex,
            ["row"] = 1,
            ["column"] = 1,
            ["text"] = "Cell A1"
        }), "set_cell_text");

        var getCellTextResult = await Call("get_cell_text", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = tableShapeIndex,
            ["row"] = 1,
            ["column"] = 1
        });
        AssertSuccess(getCellTextResult, "get_cell_text");
        Assert.Equal("Cell A1", GetString(getCellTextResult, "cellText"));
        _output.WriteLine("✓ add_table/set_cell_text, get_cell_text round-trip");

        // 6. add_chart (categories/series/values), get_chart_data.
        var addChartResult = await Call("add_chart", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["chartType"] = "bar",
            ["left"] = 20f,
            ["top"] = 320f,
            ["width"] = 300f,
            ["height"] = 150f,
            ["categories"] = ChartCategories,
            ["seriesName"] = "Revenue",
            ["values"] = ChartValues
        });
        AssertSuccess(addChartResult, "add_chart");
        var chartShapeIndex = GetInt(addChartResult, "shapeIndex")!.Value;

        var getChartDataResult = await Call("get_chart_data", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["shapeIndex"] = chartShapeIndex
        });
        AssertSuccess(getChartDataResult, "get_chart_data");
        Assert.Equal(3, GetInt(getChartDataResult, "categoryCount"));
        Assert.Equal(1, GetInt(getChartDataResult, "seriesCount"));
        _output.WriteLine("✓ add_chart, get_chart_data (3 categories, 1 series)");

        // 7. add_picture (small local PNG generated at test setup).
        var shapeCountBeforePictureResult = await Call("get_shape_count", new() { ["sessionId"] = sessionId, ["slideIndex"] = slideIndex });
        AssertSuccess(shapeCountBeforePictureResult, "get_shape_count (before picture)");
        var shapeCountBeforePicture = GetInt(shapeCountBeforePictureResult, "shapeCount")!.Value;

        var addPictureResult = await Call("add_picture", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["imagePath"] = _imageFile,
            ["left"] = 350f,
            ["top"] = 20f,
            ["width"] = 40f,
            ["height"] = 40f
        });
        AssertSuccess(addPictureResult, "add_picture");
        Assert.Equal(shapeCountBeforePicture + 1, GetInt(addPictureResult, "shapeCount"));
        _output.WriteLine("✓ add_picture");

        // 8. set_notes_text / get_notes_text round-trip.
        AssertSuccess(await Call("set_notes_text", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["text"] = "Speaker notes for slide 1."
        }), "set_notes_text");

        var getNotesResult = await Call("get_notes_text", new() { ["sessionId"] = sessionId, ["slideIndex"] = slideIndex });
        AssertSuccess(getNotesResult, "get_notes_text");
        Assert.Equal("Speaker notes for slide 1.", GetString(getNotesResult, "notesText"));
        _output.WriteLine("✓ set_notes_text/get_notes_text round-trip");

        // 9. set_layout / get_layout.
        AssertSuccess(await Call("set_layout", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["layoutName"] = "ppLayoutTitleOnly"
        }), "set_layout");

        var getLayoutResult = await Call("get_layout", new() { ["sessionId"] = sessionId, ["slideIndex"] = slideIndex });
        AssertSuccess(getLayoutResult, "get_layout");
        Assert.Equal("ppLayoutTitleOnly", GetString(getLayoutResult, "layoutName"));
        _output.WriteLine("✓ set_layout/get_layout round-trip");

        // 10. export_slide_to_image and export_all_slides_to_images.
        var singleExportPath = Path.Join(_tempDir, "slide1.png");
        var exportSlideResult = await Call("export_slide_to_image", new()
        {
            ["sessionId"] = sessionId,
            ["slideIndex"] = slideIndex,
            ["outputPath"] = singleExportPath
        });
        AssertSuccess(exportSlideResult, "export_slide_to_image");
        Assert.True(File.Exists(singleExportPath), $"Expected exported file: {singleExportPath}");
        Assert.True(new FileInfo(singleExportPath).Length > 0, "Exported single-slide image is empty.");
        _output.WriteLine($"✓ export_slide_to_image → {singleExportPath} ({new FileInfo(singleExportPath).Length} bytes)");

        var exportAllDir = Path.Join(_tempDir, "all-slides");
        var exportAllResult = await Call("export_all_slides_to_images", new()
        {
            ["sessionId"] = sessionId,
            ["outputDirectory"] = exportAllDir
        });
        AssertSuccess(exportAllResult, "export_all_slides_to_images");
        Assert.True(Directory.Exists(exportAllDir), $"Expected export directory: {exportAllDir}");
        var exportedFiles = Directory.GetFiles(exportAllDir);
        Assert.Equal(afterAddCount, exportedFiles.Length);
        Assert.All(exportedFiles, f => Assert.True(new FileInfo(f).Length > 0, $"{f} is empty"));
        _output.WriteLine($"✓ export_all_slides_to_images → {exportedFiles.Length} files in {exportAllDir}");

        // 11. save_presentation, then close_presentation (fast/non-blocking), then list_sessions
        // confirms the session is gone.
        AssertSuccess(await Call("save_presentation", new() { ["sessionId"] = sessionId }), "save_presentation");
        _output.WriteLine("✓ save_presentation");

        var closeStopwatch = System.Diagnostics.Stopwatch.StartNew();
        var closeResult = await Call("close_presentation", new() { ["sessionId"] = sessionId });
        closeStopwatch.Stop();
        AssertSuccess(closeResult, "close_presentation");
        Assert.True(GetBool(closeResult, "closed"));
        Assert.True(
            closeStopwatch.Elapsed < TimeSpan.FromSeconds(15),
            $"close_presentation should return immediately (async close per Brett's fix); took {closeStopwatch.Elapsed}.");
        _output.WriteLine($"✓ close_presentation returned in {closeStopwatch.ElapsedMilliseconds}ms (non-blocking)");

        var listSessionsResult = await Call("list_sessions", []);
        AssertSuccess(listSessionsResult, "list_sessions");
        using (var listJson = System.Text.Json.JsonDocument.Parse(listSessionsResult))
        {
            var stillFound = listJson.RootElement.GetProperty("sessions").EnumerateArray()
                .Any(s => string.Equals(s.GetProperty("sessionId").GetString(), sessionId, StringComparison.Ordinal));
            Assert.False(stillFound, $"Session {sessionId} should be gone after close_presentation: {listSessionsResult}");
        }
        _output.WriteLine("✓ list_sessions confirms the session is gone");
    }

    private Task<string> Call(string toolName, Dictionary<string, object?> arguments)
        => CallToolAsync(_client!, toolName, arguments, _cts.Token);
}
