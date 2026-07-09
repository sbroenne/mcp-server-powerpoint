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
/// Deliberately a single long test rather than many small ones: a live PowerPoint COM session is
/// expensive to spin up and tear down (the quit/grace-period sequence dominates), so paying that
/// cost more than once per test class would make the suite unreasonably slow. <c>create_presentation</c>
/// itself is now fast and non-blocking (create-and-keep-open, returns an open sessionId — see
/// .squad/decisions/inbox/brett-create-and-open.md). Every assertion here goes through a tools/call
/// JSON response — never a direct method call — per Ripley's charter (integration tests only, no
/// mocking, real COM).
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
        // 1. create_presentation returns an OPEN session (create-and-keep-open) → sessionId.
        var createResult = await Call("create_presentation", new() { ["filePath"] = _presentationFile });
        AssertSuccess(createResult, "create_presentation");
        Assert.True(File.Exists(_presentationFile));
        var sessionId = GetString(createResult, "sessionId");
        Assert.False(string.IsNullOrEmpty(sessionId));
        _output.WriteLine($"✓ create_presentation → open sessionId={sessionId}");

        // 2. slide.add-blank (x2), slide.get-count asserts the count grew by exactly 2.
        var baselineCountResult = await Call("slide", new() { ["action"] = "get-count", ["session_id"] = sessionId });
        AssertSuccess(baselineCountResult, "slide.get-count (baseline)");
        var baselineCount = GetInt(baselineCountResult, "slideCount")!.Value;

        AssertSuccess(await Call("slide", new() { ["action"] = "add-blank", ["session_id"] = sessionId }), "slide.add-blank #1");
        AssertSuccess(await Call("slide", new() { ["action"] = "add-blank", ["session_id"] = sessionId }), "slide.add-blank #2");

        var afterAddCountResult = await Call("slide", new() { ["action"] = "get-count", ["session_id"] = sessionId });
        AssertSuccess(afterAddCountResult, "slide.get-count (after add)");
        var afterAddCount = GetInt(afterAddCountResult, "slideCount")!.Value;
        Assert.Equal(baselineCount + 2, afterAddCount);
        _output.WriteLine($"✓ slide.add-blank x2, slide.get-count {baselineCount} → {afterAddCount}");

        const int slideIndex = 1;

        // 3. shape.add-text-box / textframe.set-text / set-font-size / set-bold / set-font-color; textframe.get-text round-trips.
        var addTextBoxResult = await Call("shape", new()
        {
            ["action"] = "add-text-box",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["left"] = 10f,
            ["top"] = 10f,
            ["width"] = 200f,
            ["height"] = 50f,
            ["text"] = "Hello"
        });
        AssertSuccess(addTextBoxResult, "shape.add-text-box");
        var textBoxShapeIndex = GetInt(addTextBoxResult, "shapeIndex")!.Value;
        _output.WriteLine($"✓ shape.add-text-box → shapeIndex={textBoxShapeIndex}");

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["text"] = "Updated Text"
        }), "textframe.set-text");

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-font-size",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["font_size"] = 24f
        }), "textframe.set-font-size");

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-bold",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["bold"] = true
        }), "textframe.set-bold");

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-font-color",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["red"] = (byte)255,
            ["green"] = (byte)0,
            ["blue"] = (byte)0
        }), "textframe.set-font-color");

        var getTextResult = await Call("textframe", new()
        {
            ["action"] = "get-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex
        });
        AssertSuccess(getTextResult, "textframe.get-text");
        Assert.Equal("Updated Text", GetString(getTextResult, "text"));
        _output.WriteLine("✓ textframe.set-text/set-font-size/set-bold/set-font-color, get-text round-trips");

        // 3b. textframe.set-italic/get-italic, set-underline/get-underline, set-font-name/get-font-name,
        // set-alignment/get-alignment, set-bullet/get-bullet.
        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-italic",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["italic"] = true
        }), "textframe.set-italic");
        var getItalicResult = await Call("textframe", new() { ["action"] = "get-italic", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = textBoxShapeIndex });
        AssertSuccess(getItalicResult, "textframe.get-italic");
        Assert.True(GetBool(getItalicResult, "italic"));

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-underline",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["underline"] = true
        }), "textframe.set-underline");
        var getUnderlineResult = await Call("textframe", new() { ["action"] = "get-underline", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = textBoxShapeIndex });
        AssertSuccess(getUnderlineResult, "textframe.get-underline");
        Assert.True(GetBool(getUnderlineResult, "underline"));

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-font-name",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["font_name"] = "Georgia"
        }), "textframe.set-font-name");
        var getFontNameResult = await Call("textframe", new() { ["action"] = "get-font-name", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = textBoxShapeIndex });
        AssertSuccess(getFontNameResult, "textframe.get-font-name");
        Assert.Equal("Georgia", GetString(getFontNameResult, "fontName"));

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-alignment",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["alignment"] = "ppAlignCenter"
        }), "textframe.set-alignment");
        var getAlignmentResult = await Call("textframe", new() { ["action"] = "get-alignment", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = textBoxShapeIndex });
        AssertSuccess(getAlignmentResult, "textframe.get-alignment");
        Assert.Equal("ppAlignCenter", GetString(getAlignmentResult, "alignment"));

        AssertSuccess(await Call("textframe", new()
        {
            ["action"] = "set-bullet",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["enabled"] = true,
            ["character"] = "-"
        }), "textframe.set-bullet");
        var getBulletResult = await Call("textframe", new() { ["action"] = "get-bullet", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = textBoxShapeIndex });
        AssertSuccess(getBulletResult, "textframe.get-bullet");
        Assert.True(GetBool(getBulletResult, "bulletEnabled"));
        Assert.Equal("-", GetString(getBulletResult, "bulletCharacter"));
        _output.WriteLine("✓ textframe.set-italic/set-underline/set-font-name/set-alignment/set-bullet round-trips");

        // 4. shape.add-rectangle, set-position, set-size, get-count, delete a shape.
        var addRectResult = await Call("shape", new()
        {
            ["action"] = "add-rectangle",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["left"] = 50f,
            ["top"] = 80f,
            ["width"] = 100f,
            ["height"] = 60f
        });
        AssertSuccess(addRectResult, "shape.add-rectangle");
        var rectShapeIndex = GetInt(addRectResult, "shapeIndex")!.Value;

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-position",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = rectShapeIndex,
            ["left"] = 75f,
            ["top"] = 90f
        }), "shape.set-position");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-size",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = rectShapeIndex,
            ["width"] = 120f,
            ["height"] = 70f
        }), "shape.set-size");

        var shapeCountBeforeDeleteResult = await Call("shape", new() { ["action"] = "get-count", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(shapeCountBeforeDeleteResult, "shape.get-count (before delete)");
        var shapeCountBeforeDelete = GetInt(shapeCountBeforeDeleteResult, "shapeCount")!.Value;

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "delete",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = rectShapeIndex
        }), "shape.delete");

        var shapeCountAfterDeleteResult = await Call("shape", new() { ["action"] = "get-count", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(shapeCountAfterDeleteResult, "shape.get-count (after delete)");
        var shapeCountAfterDelete = GetInt(shapeCountAfterDeleteResult, "shapeCount")!.Value;
        Assert.Equal(shapeCountBeforeDelete - 1, shapeCountAfterDelete);
        _output.WriteLine($"✓ shape.add-rectangle/set-position/set-size/delete, shapeCount {shapeCountBeforeDelete} → {shapeCountAfterDelete}");

        // 4b. shape.add-auto-shape, add-line, add-connector.
        var addAutoShapeResult = await Call("shape", new()
        {
            ["action"] = "add-auto-shape",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_type"] = "msoShapeOval",
            ["left"] = 10f,
            ["top"] = 10f,
            ["width"] = 60f,
            ["height"] = 40f
        });
        AssertSuccess(addAutoShapeResult, "shape.add-auto-shape");
        Assert.Equal("msoShapeOval", GetString(addAutoShapeResult, "shapeTypeName"));
        _output.WriteLine("✓ shape.add-auto-shape (msoShapeOval)");

        var addLineResult = await Call("shape", new()
        {
            ["action"] = "add-line",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["begin_x"] = 0f,
            ["begin_y"] = 0f,
            ["end_x"] = 100f,
            ["end_y"] = 50f
        });
        AssertSuccess(addLineResult, "shape.add-line");
        _output.WriteLine("✓ shape.add-line");

        var addConnectorResult = await Call("shape", new()
        {
            ["action"] = "add-connector",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["connector_type"] = "msoConnectorElbow",
            ["begin_x"] = 0f,
            ["begin_y"] = 0f,
            ["end_x"] = 80f,
            ["end_y"] = 80f
        });
        AssertSuccess(addConnectorResult, "shape.add-connector");
        Assert.Equal("msoConnectorElbow", GetString(addConnectorResult, "connectorTypeName"));
        _output.WriteLine("✓ shape.add-connector (msoConnectorElbow)");

        // 4c. shape formatting: fill/line/rotation/flip/z-order/shadow/group/name/alt-text.
        var formattingShapeIndex = GetInt(addAutoShapeResult, "shapeIndex")!.Value;

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-fill",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["red"] = (byte)255,
            ["green"] = (byte)0,
            ["blue"] = (byte)0
        }), "shape.set-fill");
        var getFillResult = await Call("shape", new() { ["action"] = "get-fill", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getFillResult, "shape.get-fill");
        Assert.Equal(255, GetInt(getFillResult, "colorRgb"));
        _output.WriteLine("✓ shape.set-fill/get-fill");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-line",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["weight"] = 2f,
            ["dash_style"] = "msoLineDash"
        }), "shape.set-line");
        var getLineResult = await Call("shape", new() { ["action"] = "get-line", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getLineResult, "shape.get-line");
        Assert.Equal("msoLineDash", GetString(getLineResult, "dashStyleName"));
        _output.WriteLine("✓ shape.set-line/get-line");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-rotation",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["degrees"] = 30f
        }), "shape.set-rotation");
        var getRotationResult = await Call("shape", new() { ["action"] = "get-rotation", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getRotationResult, "shape.get-rotation");
        _output.WriteLine("✓ shape.set-rotation/get-rotation");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "flip",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["direction"] = "horizontal"
        }), "shape.flip");
        _output.WriteLine("✓ shape.flip");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-z-order",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["z_order_command"] = "bring-to-front"
        }), "shape.set-z-order");
        _output.WriteLine("✓ shape.set-z-order");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-shadow",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["visible"] = true
        }), "shape.set-shadow");
        var getShadowResult = await Call("shape", new() { ["action"] = "get-shadow", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getShadowResult, "shape.get-shadow");
        Assert.True(GetBool(getShadowResult, "visible"));
        _output.WriteLine("✓ shape.set-shadow/get-shadow");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-name",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["name"] = "FormattingDemoShape"
        }), "shape.set-name");
        var getNameResult = await Call("shape", new() { ["action"] = "get-name", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getNameResult, "shape.get-name");
        Assert.Equal("FormattingDemoShape", GetString(getNameResult, "name"));
        _output.WriteLine("✓ shape.set-name/get-name");

        AssertSuccess(await Call("shape", new()
        {
            ["action"] = "set-alt-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = formattingShapeIndex,
            ["alt_text"] = "An oval used in the MCP authoring workflow test"
        }), "shape.set-alt-text");
        var getAltTextResult = await Call("shape", new() { ["action"] = "get-alt-text", ["session_id"] = sessionId, ["slide_index"] = slideIndex, ["shape_index"] = formattingShapeIndex });
        AssertSuccess(getAltTextResult, "shape.get-alt-text");
        Assert.Equal("An oval used in the MCP authoring workflow test", GetString(getAltTextResult, "altText"));
        _output.WriteLine("✓ shape.set-alt-text/get-alt-text");

        // Add two fresh rectangles at the end of the z-order so their group's resulting shape
        // index is predictable (the last index) — avoids relying on Group()'s returned dynamic
        // shape object, which is subject to the NoPIA late-binding ".Index" quirk.
        var groupRectAResult = await Call("shape", new()
        {
            ["action"] = "add-rectangle",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["left"] = 200f,
            ["top"] = 200f,
            ["width"] = 40f,
            ["height"] = 40f
        });
        AssertSuccess(groupRectAResult, "shape.add-rectangle (group A)");
        var groupRectA = GetInt(groupRectAResult, "shapeIndex")!.Value;

        var groupRectBResult = await Call("shape", new()
        {
            ["action"] = "add-rectangle",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["left"] = 250f,
            ["top"] = 200f,
            ["width"] = 40f,
            ["height"] = 40f
        });
        AssertSuccess(groupRectBResult, "shape.add-rectangle (group B)");
        var groupRectB = GetInt(groupRectBResult, "shapeIndex")!.Value;

        var groupResult = await Call("shape", new()
        {
            ["action"] = "group",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_indexes"] = new[] { groupRectA, groupRectB }
        });
        AssertSuccess(groupResult, "shape.group");
        var groupedShapeCount = GetInt(groupResult, "shapeCount")!.Value;
        _output.WriteLine($"✓ shape.group → shapeCount={groupedShapeCount}");

        // groupRectA and groupRectB were the last two shapes on the slide, so the resulting
        // group shape occupies the last index (== groupedShapeCount).
        var ungroupResult = await Call("shape", new()
        {
            ["action"] = "ungroup",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = groupedShapeCount
        });
        AssertSuccess(ungroupResult, "shape.ungroup");
        Assert.Equal(2, GetInt(ungroupResult, "ungroupedShapeCount"));
        _output.WriteLine("✓ shape.ungroup");

        // 5. table.add-table, set-cell-text, get-cell-text round-trip.
        var addTableResult = await Call("table", new()
        {
            ["action"] = "add-table",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["rows"] = 2,
            ["columns"] = 2,
            ["left"] = 20f,
            ["top"] = 200f,
            ["width"] = 300f,
            ["height"] = 100f
        });
        AssertSuccess(addTableResult, "table.add-table");
        var tableShapeIndex = GetInt(addTableResult, "shapeIndex")!.Value;

        AssertSuccess(await Call("table", new()
        {
            ["action"] = "set-cell-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = tableShapeIndex,
            ["row"] = 1,
            ["column"] = 1,
            ["text"] = "Cell A1"
        }), "table.set-cell-text");

        var getCellTextResult = await Call("table", new()
        {
            ["action"] = "get-cell-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = tableShapeIndex,
            ["row"] = 1,
            ["column"] = 1
        });
        AssertSuccess(getCellTextResult, "table.get-cell-text");
        Assert.Equal("Cell A1", GetString(getCellTextResult, "cellText"));
        _output.WriteLine("✓ table.add-table/set-cell-text, get-cell-text round-trip");

        // 6. chart.add-chart (categories/series/values), get-chart-data.
        var addChartResult = await Call("chart", new()
        {
            ["action"] = "add-chart",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["chart_type"] = "bar",
            ["left"] = 20f,
            ["top"] = 320f,
            ["width"] = 300f,
            ["height"] = 150f,
            ["categories"] = ChartCategories,
            ["series_name"] = "Revenue",
            ["values"] = ChartValues
        });
        AssertSuccess(addChartResult, "chart.add-chart");
        var chartShapeIndex = GetInt(addChartResult, "shapeIndex")!.Value;

        var getChartDataResult = await Call("chart", new()
        {
            ["action"] = "get-chart-data",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = chartShapeIndex
        });
        AssertSuccess(getChartDataResult, "chart.get-chart-data");
        Assert.Equal(3, GetInt(getChartDataResult, "categoryCount"));
        Assert.Equal(1, GetInt(getChartDataResult, "seriesCount"));
        _output.WriteLine("✓ chart.add-chart, get-chart-data (3 categories, 1 series)");

        // 7. image.add-picture (small local PNG generated at test setup).
        var shapeCountBeforePictureResult = await Call("shape", new() { ["action"] = "get-count", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(shapeCountBeforePictureResult, "shape.get-count (before picture)");
        var shapeCountBeforePicture = GetInt(shapeCountBeforePictureResult, "shapeCount")!.Value;

        var addPictureResult = await Call("image", new()
        {
            ["action"] = "add-picture",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["image_path"] = _imageFile,
            ["left"] = 350f,
            ["top"] = 20f,
            ["width"] = 40f,
            ["height"] = 40f
        });
        AssertSuccess(addPictureResult, "image.add-picture");
        Assert.Equal(shapeCountBeforePicture + 1, GetInt(addPictureResult, "shapeCount"));
        _output.WriteLine("✓ image.add-picture");

        // 8. notes.set-notes-text / get-notes-text round-trip.
        AssertSuccess(await Call("notes", new()
        {
            ["action"] = "set-notes-text",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["text"] = "Speaker notes for slide 1."
        }), "notes.set-notes-text");

        var getNotesResult = await Call("notes", new() { ["action"] = "get-notes-text", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(getNotesResult, "notes.get-notes-text");
        Assert.Equal("Speaker notes for slide 1.", GetString(getNotesResult, "notesText"));
        _output.WriteLine("✓ notes.set-notes-text/get-notes-text round-trip");

        // 9. layout.set-layout / get-layout.
        AssertSuccess(await Call("layout", new()
        {
            ["action"] = "set-layout",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["layout_name"] = "ppLayoutTitleOnly"
        }), "layout.set-layout");

        var getLayoutResult = await Call("layout", new() { ["action"] = "get-layout", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(getLayoutResult, "layout.get-layout");
        Assert.Equal("ppLayoutTitleOnly", GetString(getLayoutResult, "layoutName"));
        _output.WriteLine("✓ layout.set-layout/get-layout round-trip");

        // 10. master.set-title-font / get-title-font, set-background-color / get-background-color.
        AssertSuccess(await Call("master", new()
        {
            ["action"] = "set-title-font",
            ["session_id"] = sessionId,
            ["font_name"] = "Georgia",
            ["font_size"] = 36.0,
            ["bold"] = true
        }), "master.set-title-font");

        var getTitleFontResult = await Call("master", new() { ["action"] = "get-title-font", ["session_id"] = sessionId });
        AssertSuccess(getTitleFontResult, "master.get-title-font");
        Assert.Equal("Georgia", GetString(getTitleFontResult, "fontName"));
        Assert.True(GetBool(getTitleFontResult, "bold"));
        _output.WriteLine("✓ master.set-title-font/get-title-font round-trip");

        AssertSuccess(await Call("master", new()
        {
            ["action"] = "set-background-color",
            ["session_id"] = sessionId,
            ["red"] = 240,
            ["green"] = 240,
            ["blue"] = 240
        }), "master.set-background-color");

        var getBackgroundResult = await Call("master", new() { ["action"] = "get-background-color", ["session_id"] = sessionId });
        AssertSuccess(getBackgroundResult, "master.get-background-color");
        _output.WriteLine("✓ master.set-background-color/get-background-color round-trip");

        // 10b. animation.add-effect (entrance + exit), get-effect-count, delete-effect;
        // animation.set-transition / get-transition round-trip.
        var addEffectResult = await Call("animation", new()
        {
            ["action"] = "add-effect",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["effect_name"] = "msoAnimEffectFade",
            ["trigger"] = "with-previous"
        });
        AssertSuccess(addEffectResult, "animation.add-effect (entrance)");
        Assert.Equal("msoAnimEffectFade", GetString(addEffectResult, "effectName"));
        Assert.False(GetBool(addEffectResult, "isExit"));

        var addExitEffectResult = await Call("animation", new()
        {
            ["action"] = "add-effect",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["shape_index"] = textBoxShapeIndex,
            ["effect_name"] = "msoAnimEffectFly",
            ["is_exit"] = true,
            ["trigger"] = "after-previous"
        });
        AssertSuccess(addExitEffectResult, "animation.add-effect (exit)");
        Assert.True(GetBool(addExitEffectResult, "isExit"));
        var exitEffectIndex = GetInt(addExitEffectResult, "effectIndex")!.Value;

        var effectCountResult = await Call("animation", new() { ["action"] = "get-effect-count", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(effectCountResult, "animation.get-effect-count");
        Assert.Equal(2, GetInt(effectCountResult, "effectCount"));

        AssertSuccess(await Call("animation", new()
        {
            ["action"] = "delete-effect",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["effect_index"] = exitEffectIndex
        }), "animation.delete-effect");

        var effectCountAfterDeleteResult = await Call("animation", new() { ["action"] = "get-effect-count", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(effectCountAfterDeleteResult, "animation.get-effect-count (after delete)");
        Assert.Equal(1, GetInt(effectCountAfterDeleteResult, "effectCount"));
        _output.WriteLine("✓ animation.add-effect (entrance+exit)/get-effect-count/delete-effect");

        AssertSuccess(await Call("animation", new()
        {
            ["action"] = "set-transition",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["transition_name"] = "ppEffectFade",
            ["duration_seconds"] = 1.25
        }), "animation.set-transition");

        var getTransitionResult = await Call("animation", new() { ["action"] = "get-transition", ["session_id"] = sessionId, ["slide_index"] = slideIndex });
        AssertSuccess(getTransitionResult, "animation.get-transition");
        Assert.Equal("ppEffectFade", GetString(getTransitionResult, "transitionName"));
        _output.WriteLine("✓ animation.set-transition/get-transition round-trip");

        // 11. export.export-slide-to-image and export-all-slides-to-images.
        var singleExportPath = Path.Join(_tempDir, "slide1.png");
        var exportSlideResult = await Call("export", new()
        {
            ["action"] = "export-slide-to-image",
            ["session_id"] = sessionId,
            ["slide_index"] = slideIndex,
            ["output_path"] = singleExportPath
        });
        AssertSuccess(exportSlideResult, "export.export-slide-to-image");
        Assert.True(File.Exists(singleExportPath), $"Expected exported file: {singleExportPath}");
        Assert.True(new FileInfo(singleExportPath).Length > 0, "Exported single-slide image is empty.");
        _output.WriteLine($"✓ export.export-slide-to-image → {singleExportPath} ({new FileInfo(singleExportPath).Length} bytes)");

        var exportAllDir = Path.Join(_tempDir, "all-slides");
        var exportAllResult = await Call("export", new()
        {
            ["action"] = "export-all-slides-to-images",
            ["session_id"] = sessionId,
            ["output_directory"] = exportAllDir
        });
        AssertSuccess(exportAllResult, "export.export-all-slides-to-images");
        Assert.True(Directory.Exists(exportAllDir), $"Expected export directory: {exportAllDir}");
        var exportedFiles = Directory.GetFiles(exportAllDir);
        Assert.Equal(afterAddCount, exportedFiles.Length);
        Assert.All(exportedFiles, f => Assert.True(new FileInfo(f).Length > 0, $"{f} is empty"));
        _output.WriteLine($"✓ export.export-all-slides-to-images → {exportedFiles.Length} files in {exportAllDir}");

        // 12. save_presentation, then close_presentation (fast/non-blocking), then list_sessions
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
