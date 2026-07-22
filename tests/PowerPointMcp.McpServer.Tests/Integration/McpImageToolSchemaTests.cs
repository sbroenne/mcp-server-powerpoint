// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Client;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Generated;
using Sbroenne.PowerPointMcp.McpServer.Tools;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Future-proof schema tests that verify every <c>IImageCommands</c> method is correctly
/// translated to the generated MCP <c>image</c> tool via <c>tools/list</c>, AND that the
/// generated CLI surface (via <see cref="ServiceRegistry.Image.CliSettings"/> and
/// <see cref="ServiceRegistry.Image.RouteCliArgs"/>) exposes every action and parameter.
/// Tests fail if any interface method is omitted, renamed, or its parameters are mistranslated
/// by the generator. No PowerPoint COM instance is required — MCP tests run via in-memory
/// transport; CLI tests are no-transport reflective assertions.
/// </summary>
/// <remarks>
/// Design intent: ground-truth is derived from <see cref="ServiceRegistry.Image.ValidActions"/>
/// (itself generated from <c>IImageCommands</c> by <c>ServiceRegistryGenerator</c>) and the
/// stable camelCase → snake_case naming rule, so adding a new method to <c>IImageCommands</c>
/// automatically propagates into an assertion failure here until both the MCP and CLI surfaces
/// catch up. The <see cref="ImageAction"/> enum cross-check (no transport needed) additionally
/// guards that the enum members stay in sync with <c>ValidActions</c>.
/// </remarks>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Image")]
public sealed class McpImageToolSchemaTests : IAsyncLifetime, IAsyncDisposable
{
    // ─────────────────────────────────────────────────────────────────────────────
    // Ground truth: derived from ServiceRegistry (generated from IImageCommands).
    // Adding a method to IImageCommands (and re-running the generator) extends ValidActions,
    // which immediately makes the protocol assertions below fail until the MCP surface
    // is regenerated to include the new action and its parameters.
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Expected MCP parameter names: fixed generator params + the union of all
    /// <c>IImageCommands</c> method parameters (excluding <c>IPresentationBatch</c>),
    /// converted via the stable camelCase → snake_case naming rule (e.g. <c>slideIndex</c> →
    /// <c>slide_index</c>). See the inline comments on <see cref="ExpectedMcpParameters"/> below
    /// for the per-parameter interface name and which action(s) require it.
    /// </summary>
    private static readonly HashSet<string> ExpectedMcpParameters = new(StringComparer.Ordinal)
    {
        "action",       // generated: ImageAction enum, one entry per IImageCommands method
        "session_id",   // generator fixed: added for all session-aware tools
        "slide_index",  // slideIndex  → snake_case (required for all 7 actions)
        "image_path",   // imagePath   → snake_case (required for: add-picture)
        "left",         // left        → unchanged  (required for: add-picture)
        "top",          // top         → unchanged  (required for: add-picture)
        "width",        // width       → unchanged  (required for: add-picture)
        "height",       // height      → unchanged  (required for: add-picture)
        "shape_index",  // shapeIndex  → snake_case (required for: set/get-brightness-contrast, set/get-recolor, set/get-crop)
        "brightness",   // brightness  → unchanged  (required for: set-brightness-contrast)
        "contrast",     // contrast    → unchanged  (required for: set-brightness-contrast)
        "color_type",   // colorType   → snake_case (required for: set-recolor)
        "crop_left",    // cropLeft    → snake_case (required for: set-crop)
        "crop_top",     // cropTop     → snake_case (required for: set-crop)
        "crop_right",   // cropRight   → snake_case (required for: set-crop)
        "crop_bottom",  // cropBottom  → snake_case (required for: set-crop)
    };

    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpImageToolSchemaTests(ITestOutputHelper output) => _output = output;

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "ImageSchemaTestClient",
            _cts.Token);

        _output.WriteLine($"✓ Connected: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");
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
            _client, _clientToServerPipe, _serverToClientPipe, _serverTask, _output);
        _cts.Dispose();
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // No-transport guard: ImageAction enum ↔ ValidActions in sync
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The generated <see cref="ImageAction"/> enum must have exactly as many members as
    /// <see cref="ServiceRegistry.Image.ValidActions"/> and every action must round-trip through
    /// <see cref="ServiceRegistry.Image.TryParseAction"/>. Both are generated from
    /// <c>IImageCommands</c> — this cross-check guards that the two generated artifacts stay in
    /// sync. Does not require a transport connection.
    /// </summary>
    [Fact]
    public void ImageAction_EnumAndValidActions_AreInSync()
    {
        var enumActions = Enum.GetValues<ImageAction>()
            .Select(a => ServiceRegistry.Image.ToActionString(a))
            .ToList();

        Assert.Equal(
            ServiceRegistry.Image.ValidActions.Length,
            enumActions.Count);

        // Every ValidAction must parse to an enum member and convert back losslessly.
        foreach (var action in ServiceRegistry.Image.ValidActions)
        {
            Assert.True(
                ServiceRegistry.Image.TryParseAction(action, out var parsed),
                $"ValidAction '{action}' (from IImageCommands) cannot be parsed to an ImageAction enum member.");

            var roundTripped = ServiceRegistry.Image.ToActionString(parsed);
            Assert.True(
                action == roundTripped,
                $"Action '{action}' did not survive a round-trip through the enum: got '{roundTripped}'.");
        }

        _output.WriteLine(
            $"✓ ImageAction enum ({enumActions.Count} members) matches ValidActions and round-trips correctly");
    }

    [Fact]
    public void ImageCropResult_SerializesResponseFieldsAsCamelCase()
    {
        var result = new ImageOperationResult
        {
            Success = true,
            CropLeft = 1f,
            CropTop = 2f,
            CropRight = 3f,
            CropBottom = 4f
        };

        using var document = JsonDocument.Parse(PowerPointToolsBase.Serialize(result));
        JsonElement root = document.RootElement;

        Assert.Equal(1f, root.GetProperty("cropLeft").GetSingle());
        Assert.Equal(2f, root.GetProperty("cropTop").GetSingle());
        Assert.Equal(3f, root.GetProperty("cropRight").GetSingle());
        Assert.Equal(4f, root.GetProperty("cropBottom").GetSingle());
        Assert.False(root.TryGetProperty("crop_left", out _));
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // Protocol assertions: tools/list schema (no COM)
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The top-level description of the MCP <c>image</c> tool must mention every action string
    /// from <see cref="ServiceRegistry.Image.ValidActions"/>. The description is generated from
    /// the list of interface methods, so dropping a method from <c>IImageCommands</c> (and
    /// regenerating) causes the description to shrink and this test to fail.
    /// </summary>
    [Fact]
    public async Task ImageTool_ToolDescription_MentionsEveryInterfaceAction()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var imageTool = tools.SingleOrDefault(t => t.Name == ServiceRegistry.Image.McpToolName);
        Assert.NotNull(imageTool);

        var description = imageTool.Description ?? "";
        _output.WriteLine($"Tool description: {description}");

        foreach (var action in ServiceRegistry.Image.ValidActions)
        {
            Assert.True(
                description.Contains(action, StringComparison.Ordinal),
                $"Image tool description is missing action '{action}' (from IImageCommands.{action}). " +
                $"The generator may have dropped a method from the interface.");
        }

        _output.WriteLine(
            $"✓ Tool description mentions all {ServiceRegistry.Image.ValidActions.Length} IImageCommands actions");
    }

    /// <summary>
    /// The <c>action</c> parameter description in the MCP schema must mention every action from
    /// <see cref="ServiceRegistry.Image.ValidActions"/>. This is separate from the top-level tool
    /// description so both description surfaces are covered.
    /// </summary>
    [Fact]
    public async Task ImageTool_ActionParameterDescription_MentionsEveryInterfaceAction()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var imageTool = tools.Single(t => t.Name == ServiceRegistry.Image.McpToolName);

        var schema = imageTool.JsonSchema;
        Assert.True(
            schema.TryGetProperty("properties", out var properties),
            "Image tool schema is missing the 'properties' object.");

        var actionDesc = GetPropertyDescription(properties, "action");
        _output.WriteLine($"action parameter description: {actionDesc}");

        foreach (var action in ServiceRegistry.Image.ValidActions)
        {
            Assert.True(
                actionDesc.Contains(action, StringComparison.Ordinal),
                $"Action parameter description is missing '{action}' (from IImageCommands). " +
                $"The generator may have dropped a method from the 'action' enum description.");
        }

        _output.WriteLine(
            $"✓ 'action' parameter description mentions all {ServiceRegistry.Image.ValidActions.Length} actions");
    }

    /// <summary>
    /// The <c>image</c> tool's JSON schema must expose exactly the parameters from
    /// <see cref="ExpectedMcpParameters"/> — no more, no less. Fails if a parameter is omitted,
    /// renamed (mis-converted from camelCase to snake_case), or an unexpected extra parameter
    /// appears in the generated output.
    /// </summary>
    [Fact]
    public async Task ImageTool_Schema_ContainsExactlyAllInterfaceParameters()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var imageTool = tools.Single(t => t.Name == ServiceRegistry.Image.McpToolName);

        var schema = imageTool.JsonSchema;
        Assert.True(
            schema.TryGetProperty("properties", out var properties),
            "Image tool schema is missing the 'properties' object.");

        var actualParams = properties.EnumerateObject()
            .Select(p => p.Name)
            .ToHashSet(StringComparer.Ordinal);

        _output.WriteLine(
            $"Actual schema params ({actualParams.Count}): " +
            $"{string.Join(", ", actualParams.OrderBy(x => x))}");
        _output.WriteLine(
            $"Expected schema params ({ExpectedMcpParameters.Count}): " +
            $"{string.Join(", ", ExpectedMcpParameters.OrderBy(x => x))}");

        var missing = ExpectedMcpParameters.Except(actualParams).OrderBy(x => x).ToList();
        var extra = actualParams.Except(ExpectedMcpParameters).OrderBy(x => x).ToList();

        Assert.True(
            missing.Count == 0,
            $"Image tool schema is missing parameters that IImageCommands declares: " +
            $"{string.Join(", ", missing)}. " +
            $"Either the generator dropped a method/parameter or the snake_case conversion changed.");

        Assert.True(
            extra.Count == 0,
            $"Image tool schema has unexpected extra parameters not in IImageCommands: " +
            $"{string.Join(", ", extra)}. Update ExpectedMcpParameters if IImageCommands was intentionally extended.");

        _output.WriteLine(
            $"✓ Image tool schema has exactly {actualParams.Count} expected parameters");
    }

    /// <summary>
    /// The parameter descriptions in the <c>image</c> tool schema must correctly document
    /// which actions require each parameter — matching <c>IImageCommands</c> method signatures.
    /// Changing which parameters a method accepts (e.g., adding a new parameter to
    /// <c>SetRecolor</c>) should cause the "required for" description to change and fail here.
    /// </summary>
    [Fact]
    public async Task ImageTool_ParameterDescriptions_CorrectlyDocumentRequiredByActions()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var imageTool = tools.Single(t => t.Name == ServiceRegistry.Image.McpToolName);

        var schema = imageTool.JsonSchema;
        schema.TryGetProperty("properties", out var properties);

        // image_path: required only for add-picture
        AssertRequiredFor(properties, "image_path",
            required: [ServiceRegistry.Image.AddPictureAction],
            notRequired: [
                ServiceRegistry.Image.SetBrightnessContrastAction,
                ServiceRegistry.Image.GetBrightnessContrastAction,
                ServiceRegistry.Image.SetRecolorAction,
                ServiceRegistry.Image.GetRecolorAction,
                ServiceRegistry.Image.SetCropAction,
                ServiceRegistry.Image.GetCropAction]);

        // color_type: required only for set-recolor
        AssertRequiredFor(properties, "color_type",
            required: [ServiceRegistry.Image.SetRecolorAction],
            notRequired: [
                ServiceRegistry.Image.AddPictureAction,
                ServiceRegistry.Image.SetBrightnessContrastAction,
                ServiceRegistry.Image.GetBrightnessContrastAction,
                ServiceRegistry.Image.GetRecolorAction,
                ServiceRegistry.Image.SetCropAction,
                ServiceRegistry.Image.GetCropAction]);

        // brightness + contrast: required only for set-brightness-contrast
        AssertRequiredFor(properties, "brightness",
            required: [ServiceRegistry.Image.SetBrightnessContrastAction],
            notRequired: [
                ServiceRegistry.Image.AddPictureAction,
                ServiceRegistry.Image.SetRecolorAction,
                ServiceRegistry.Image.GetRecolorAction,
                ServiceRegistry.Image.SetCropAction,
                ServiceRegistry.Image.GetCropAction]);

        AssertRequiredFor(properties, "contrast",
            required: [ServiceRegistry.Image.SetBrightnessContrastAction],
            notRequired: [
                ServiceRegistry.Image.AddPictureAction,
                ServiceRegistry.Image.SetRecolorAction,
                ServiceRegistry.Image.GetRecolorAction,
                ServiceRegistry.Image.SetCropAction,
                ServiceRegistry.Image.GetCropAction]);

        // shape_index: required for every action except add-picture
        AssertRequiredFor(properties, "shape_index",
            required: [
                ServiceRegistry.Image.SetBrightnessContrastAction,
                ServiceRegistry.Image.GetBrightnessContrastAction,
                ServiceRegistry.Image.SetRecolorAction,
                ServiceRegistry.Image.GetRecolorAction,
                ServiceRegistry.Image.SetCropAction,
                ServiceRegistry.Image.GetCropAction],
            notRequired: [ServiceRegistry.Image.AddPictureAction]);

        // crop_left / crop_top / crop_right / crop_bottom: required only for set-crop
        var cropNotRequired = new[]
        {
            ServiceRegistry.Image.AddPictureAction,
            ServiceRegistry.Image.SetBrightnessContrastAction,
            ServiceRegistry.Image.GetBrightnessContrastAction,
            ServiceRegistry.Image.SetRecolorAction,
            ServiceRegistry.Image.GetRecolorAction,
            ServiceRegistry.Image.GetCropAction,
        };

        AssertRequiredFor(properties, "crop_left",
            required: [ServiceRegistry.Image.SetCropAction],
            notRequired: cropNotRequired);

        AssertRequiredFor(properties, "crop_top",
            required: [ServiceRegistry.Image.SetCropAction],
            notRequired: cropNotRequired);

        AssertRequiredFor(properties, "crop_right",
            required: [ServiceRegistry.Image.SetCropAction],
            notRequired: cropNotRequired);

        AssertRequiredFor(properties, "crop_bottom",
            required: [ServiceRegistry.Image.SetCropAction],
            notRequired: cropNotRequired);

        _output.WriteLine(
            "✓ All parameter descriptions correctly document required-by-action constraints");
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // CLI surface assertions (no transport, no COM)
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The generated <see cref="ServiceRegistry.Image.CliSettings"/> class must expose exactly the
    /// expected CLI option long-names derived from <c>IImageCommands</c> parameter names via the
    /// stable camelCase → kebab-case naming rule. Verified via reflection on
    /// <c>CommandOptionAttribute</c> templates — fails if a parameter is dropped, renamed, or added
    /// unexpectedly by the generator. Does not require a transport connection or PowerPoint.
    /// </summary>
    [Fact]
    public void ImageCli_CliSettings_ContainsExactlyExpectedOptions()
    {
        // Extract --long-names from [CommandOption] attribute templates on CliSettings properties.
        // Template format: "-s|--session <SESSION>" or "--crop-left <CROPLEFT>"
        var cliOptionLongNames = typeof(ServiceRegistry.Image.CliSettings)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SelectMany(p => p.GetCustomAttributesData())
            .Where(a => a.AttributeType.Name == "CommandOptionAttribute")
            .Select(a => a.ConstructorArguments.FirstOrDefault().Value?.ToString() ?? "")
            // Take the first whitespace-delimited token (option flags, before the meta-var)
            .Select(template => template.Split(' ', 2)[0])
            // Split on | for "-s|--session" combined forms
            .SelectMany(token => token.Split('|'))
            .Where(part => part.StartsWith("--", StringComparison.Ordinal))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        // Expected long-names: session (fixed) + one per IImageCommands param (camelCase → kebab)
        // plus output (fixed). Derived from IImageCommands method signatures via the generator's
        // stable camelCase→kebab rule: slideIndex→--slide-index, cropLeft→--crop-left, etc.
        var expectedOptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "--session",       // fixed generator param (-s|--session)
            "--slide-index",   // slideIndex  → kebab (all 7 actions)
            "--image-path",    // imagePath   → kebab (add-picture)
            "--left",          // left        → unchanged (add-picture)
            "--top",           // top         → unchanged (add-picture)
            "--width",         // width       → unchanged (add-picture)
            "--height",        // height      → unchanged (add-picture)
            "--shape-index",   // shapeIndex  → kebab (set/get-brightness-contrast, set/get-recolor, set/get-crop)
            "--brightness",    // brightness  → unchanged (set-brightness-contrast)
            "--contrast",      // contrast    → unchanged (set-brightness-contrast)
            "--color-type",    // colorType   → kebab (set-recolor)
            "--crop-left",     // cropLeft    → kebab (set-crop)
            "--crop-top",      // cropTop     → kebab (set-crop)
            "--crop-right",    // cropRight   → kebab (set-crop)
            "--crop-bottom",   // cropBottom  → kebab (set-crop)
            "--output",        // fixed generator param (-o|--output)
        };

        _output.WriteLine(
            $"CLI option long-names ({cliOptionLongNames.Count}): " +
            $"{string.Join(", ", cliOptionLongNames.OrderBy(x => x))}");
        _output.WriteLine(
            $"Expected options ({expectedOptions.Count}): " +
            $"{string.Join(", ", expectedOptions.OrderBy(x => x))}");

        var missing = expectedOptions.Except(cliOptionLongNames).OrderBy(x => x).ToList();
        var extra = cliOptionLongNames.Except(expectedOptions).OrderBy(x => x).ToList();

        Assert.True(
            missing.Count == 0,
            $"CLI CliSettings is missing options for IImageCommands parameters: " +
            $"{string.Join(", ", missing)}. The generator may have dropped a parameter or the camelCase→kebab conversion changed.");

        Assert.True(
            extra.Count == 0,
            $"CLI CliSettings has unexpected extra options not in IImageCommands: " +
            $"{string.Join(", ", extra)}. Update expectedOptions if IImageCommands was intentionally extended.");

        _output.WriteLine(
            $"✓ CLI CliSettings exposes exactly {cliOptionLongNames.Count} expected option long-names");
    }

    /// <summary>
    /// <see cref="ServiceRegistry.Image.RouteCliArgs"/> must dispatch all 7
    /// <c>IImageCommands</c> actions and return the correct command string for each.
    /// Verifies that <c>set-crop</c> and <c>get-crop</c> are properly wired in the generated CLI
    /// dispatch — a generator omission would throw <see cref="ArgumentException"/> here.
    /// Does not require a transport connection or PowerPoint.
    /// </summary>
    [Fact]
    public void ImageCli_RouteCliArgs_DispatchesAllSevenActions()
    {
        // add-picture (imagePath required — non-empty enforced by generated RequireNotEmpty guard)
        var (addPicture, _) = ServiceRegistry.Image.RouteCliArgs(
            "add-picture", slideIndex: 1, imagePath: "C:\\test.png", left: 0, top: 0, width: 100, height: 100);
        Assert.Equal("image.add-picture", addPicture);

        // set-brightness-contrast
        var (setBC, _) = ServiceRegistry.Image.RouteCliArgs(
            "set-brightness-contrast", slideIndex: 1, shapeIndex: 1, brightness: 0.5f, contrast: 0.5f);
        Assert.Equal("image.set-brightness-contrast", setBC);

        // get-brightness-contrast
        var (getBC, _) = ServiceRegistry.Image.RouteCliArgs(
            "get-brightness-contrast", slideIndex: 1, shapeIndex: 1);
        Assert.Equal("image.get-brightness-contrast", getBC);

        // set-recolor (colorType required — non-empty enforced by generated RequireNotEmpty guard)
        var (setRecolor, _) = ServiceRegistry.Image.RouteCliArgs(
            "set-recolor", slideIndex: 1, shapeIndex: 1, colorType: "msoPictureAutomatic");
        Assert.Equal("image.set-recolor", setRecolor);

        // get-recolor
        var (getRecolor, _) = ServiceRegistry.Image.RouteCliArgs(
            "get-recolor", slideIndex: 1, shapeIndex: 1);
        Assert.Equal("image.get-recolor", getRecolor);

        // set-crop — all four crop params wired through to the dispatch
        var (setCrop, _) = ServiceRegistry.Image.RouteCliArgs(
            "set-crop", slideIndex: 1, shapeIndex: 1,
            cropLeft: 10f, cropTop: 20f, cropRight: 30f, cropBottom: 40f);
        Assert.Equal("image.set-crop", setCrop);

        // get-crop
        var (getCrop, _) = ServiceRegistry.Image.RouteCliArgs(
            "get-crop", slideIndex: 1, shapeIndex: 1);
        Assert.Equal("image.get-crop", getCrop);

        _output.WriteLine("✓ RouteCliArgs dispatches all 7 IImageCommands actions with correct command strings");
    }

    /// <summary>
    /// <see cref="ServiceRegistry.Image.ValidActions"/> must contain exactly the 7 action strings
    /// derived from <c>IImageCommands</c>, and every action must be dispatchable via
    /// <see cref="ServiceRegistry.Image.RouteCliArgs"/>. Guards against both omissions (a new
    /// method added to the interface but not to ValidActions) and accidental additions.
    /// Does not require a transport connection or PowerPoint.
    /// </summary>
    [Fact]
    public void ImageCli_ValidActions_ExactlyMatchInterfaceMethods()
    {
        var expectedActions = new HashSet<string>(StringComparer.Ordinal)
        {
            "add-picture",
            "set-brightness-contrast",
            "get-brightness-contrast",
            "set-recolor",
            "get-recolor",
            "set-crop",
            "get-crop",
        };

        var actualActions = new HashSet<string>(
            ServiceRegistry.Image.ValidActions, StringComparer.Ordinal);

        _output.WriteLine(
            $"ValidActions ({actualActions.Count}): {string.Join(", ", actualActions.OrderBy(x => x))}");

        var missing = expectedActions.Except(actualActions).OrderBy(x => x).ToList();
        var extra = actualActions.Except(expectedActions).OrderBy(x => x).ToList();

        Assert.True(
            missing.Count == 0,
            $"ValidActions is missing IImageCommands methods: {string.Join(", ", missing)}");

        Assert.True(
            extra.Count == 0,
            $"ValidActions has unexpected extra actions: {string.Join(", ", extra)}. " +
            $"Update expectedActions if IImageCommands was intentionally extended.");

        _output.WriteLine(
            $"✓ ValidActions contains exactly {actualActions.Count} expected actions");
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────────

    private void AssertRequiredFor(
        JsonElement properties,
        string paramName,
        string[] required,
        string[] notRequired)
    {
        var desc = GetPropertyDescription(properties, paramName);
        _output.WriteLine($"  {paramName}: '{desc}'");

        foreach (var action in required)
        {
            Assert.True(
                desc.Contains(action, StringComparison.Ordinal),
                $"Parameter '{paramName}' description should mention '{action}' (IImageCommands says it is required) but does not. " +
                $"Description: '{desc}'");
        }

        foreach (var action in notRequired)
        {
            Assert.False(
                desc.Contains(action, StringComparison.Ordinal),
                $"Parameter '{paramName}' description should NOT mention '{action}' (not required by that action) but does. " +
                $"Description: '{desc}'");
        }
    }

    private static string GetPropertyDescription(JsonElement properties, string name)
    {
        if (!properties.TryGetProperty(name, out var prop))
            return string.Empty;
        if (prop.TryGetProperty("description", out var desc))
            return desc.GetString() ?? string.Empty;
        return string.Empty;
    }
}
