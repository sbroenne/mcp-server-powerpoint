using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.Generated;
using Sbroenne.PowerPointMcp.McpServer.Tools;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "GeneratedContract")]
public sealed class GeneratedContractTests
{
    public enum ContractValue
    {
        FirstValue,
        SecondValue
    }

    public sealed class EnumArgs
    {
        public ContractValue Value { get; set; }
    }

    [Fact]
    public void RouteCliArgs_RejectsParameterThatDoesNotApplyToAction()
    {
        var error = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Image.RouteCliArgs(
                "add-picture",
                slideIndex: 1,
                imagePath: "image.png",
                left: 0,
                top: 0,
                width: 100,
                height: 100,
                brightness: 0.5f));

        Assert.Contains("brightness", error.Message, StringComparison.Ordinal);
        Assert.Contains("add-picture", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RouteAction_CanReturnPendingAsyncForwardResult()
    {
        var routeMethod = typeof(ServiceRegistry.Slide).GetMethod("RouteAction");
        Assert.NotNull(routeMethod);
        Assert.True(routeMethod.IsGenericMethodDefinition);
        Assert.True(ServiceRegistry.Slide.TryParseAction("get-count", out var action));

        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        Func<string, string, object?, Task<string>> forward = (_, _, _) => completion.Task;
        var arguments = routeMethod.GetParameters()
            .Select(parameter => parameter.Name switch
            {
                "action" => (object)action,
                "sessionId" => "test-session",
                "forwardToService" => forward,
                _ => parameter.DefaultValue
            })
            .ToArray();

        var result = Assert.IsAssignableFrom<Task<string>>(
            routeMethod.MakeGenericMethod(typeof(Task<string>)).Invoke(null, arguments));
        Assert.False(result.IsCompleted);

        completion.SetResult("forwarded");
        Assert.Equal("forwarded", await result);
    }

    [Theory]
    [InlineData("""{"slideIndex":1,"imagePath":"image.png","left":0,"top":0,"width":100,"height":100,"unexpected":true}""")]
    [InlineData("""{"SlideIndex":1,"imagePath":"image.png","left":0,"top":0,"width":100,"height":100}""")]
    public void DeserializeArgs_RejectsUnknownOrMisCasedProperties(string json)
    {
        Assert.Throws<System.Text.Json.JsonException>(
            () => ServiceRegistry.DeserializeArgs<ServiceRegistry.Image.AddPictureArgs>(json));
    }

    [Fact]
    public void DispatchToCore_RejectsMissingRequiredParameterBeforeCoreCall()
    {
        var error = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Image.ValidateActionArguments(
                "add-picture",
                """{"slideIndex":1,"left":0,"top":0,"width":100,"height":100}"""));

        Assert.Contains("imagePath", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AllowEmptyString_StillRequiresPropertyButAcceptsEmptyText()
    {
        ServiceRegistry.TextFrame.ValidateActionArguments(
            "set-text",
            """{"slideIndex":1,"shapeIndex":1,"text":""}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.TextFrame.ValidateActionArguments(
                "set-text",
                """{"slideIndex":1,"shapeIndex":1}"""));
    }

    [Fact]
    public void DeserializeArgs_RejectsNumericEnums()
    {
        Assert.Throws<System.Text.Json.JsonException>(
            () => ServiceRegistry.DeserializeArgs<EnumArgs>("""{"value":1}"""));
    }

    [Fact]
    public void ParseEnumValue_RejectsUnknownValuesAndAcceptsAliases()
    {
        Assert.Equal(
            ContractValue.SecondValue,
            ServiceRegistry.ParseEnumValue(
                "second-value",
                ContractValue.FirstValue,
                "value"));
        Assert.Equal(
            ContractValue.SecondValue,
            ServiceRegistry.ParseEnumValue(
                "legacy",
                ContractValue.FirstValue,
                "value",
                ("legacy", ContractValue.SecondValue)));

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ParseEnumValue(
                "unknown",
                ContractValue.FirstValue,
                "value"));
    }

    [Fact]
    public void TagActions_AreGeneratedForSlideAndShape()
    {
        string[] expected = ["set-tag", "get-tag", "list-tags", "delete-tag"];

        Assert.All(expected, action => Assert.Contains(action, ServiceRegistry.Slide.ValidActions));
        Assert.All(expected, action => Assert.Contains(action, ServiceRegistry.Shape.ValidActions));
    }

    [Fact]
    public void TagActions_EnforceRequiredAndApplicableParameters()
    {
        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Slide.ValidateActionArguments(
                "set-tag",
                """{"slideIndex":1,"tagName":"OWNER"}"""));

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.RouteCliArgs(
                "get-tag",
                slideIndex: 1,
                shapeIndex: 1,
                tagName: "OWNER",
                tagValue: "not-applicable"));
    }

    [Fact]
    public void SlideVisibilityActions_HaveGeneratedCliAndServiceWiring()
    {
        Assert.Contains("set-hidden", ServiceRegistry.Slide.ValidActions);
        Assert.Contains("set-display-master-shapes", ServiceRegistry.Slide.ValidActions);

        Assert.Equal(
            "slide.set-hidden",
            ServiceRegistry.Slide.RouteCliArgs("set-hidden", slideIndex: 1, hidden: true).Command);
        Assert.Equal(
            "slide.set-display-master-shapes",
            ServiceRegistry.Slide.RouteCliArgs("set-display-master-shapes", slideIndex: 1, display: false).Command);

        ServiceRegistry.Slide.ValidateActionArguments(
            "set-hidden",
            """{"slideIndex":1,"hidden":true}""");
        ServiceRegistry.Slide.ValidateActionArguments(
            "set-display-master-shapes",
            """{"slideIndex":1,"display":false}""");
        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Slide.ValidateActionArguments("set-hidden", """{"slideIndex":1}"""));
        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Slide.ValidateActionArguments("set-display-master-shapes", """{"slideIndex":1}"""));
    }

    [Fact]
    public void SlideVisibilityResult_SerializesReturnedStates()
    {
        var result = new SlideOperationResult
        {
            Success = true,
            Hidden = true,
            DisplaysMasterShapes = false
        };

        using var document = JsonDocument.Parse(PowerPointToolsBase.Serialize(result));
        Assert.True(document.RootElement.GetProperty("hidden").GetBoolean());
        Assert.False(document.RootElement.GetProperty("displaysMasterShapes").GetBoolean());
    }

    [Fact]
    public void ChartCli_RoutesQuickFormattingActions()
    {
        string[] expectedActions =
        [
            "get-style",
            "set-style",
            "get-color-style",
            "set-color-style",
            "get-data-table",
            "set-data-table"
        ];

        foreach (string action in expectedActions)
        {
            Assert.Contains(action, ServiceRegistry.Chart.ValidActions);
        }

        Assert.Equal(
            "chart.get-style",
            ServiceRegistry.Chart.RouteCliArgs("get-style", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "chart.set-style",
            ServiceRegistry.Chart.RouteCliArgs("set-style", slideIndex: 1, shapeIndex: 1, style: 2).Command);
        Assert.Equal(
            "chart.get-color-style",
            ServiceRegistry.Chart.RouteCliArgs("get-color-style", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "chart.set-color-style",
            ServiceRegistry.Chart.RouteCliArgs("set-color-style", slideIndex: 1, shapeIndex: 1, colorStyle: 2).Command);
        Assert.Equal(
            "chart.get-data-table",
            ServiceRegistry.Chart.RouteCliArgs("get-data-table", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "chart.set-legend-visibility",
            ServiceRegistry.Chart.RouteCliArgs("set-legend-visibility", slideIndex: 1, shapeIndex: 1, visible: true).Command);
        Assert.Equal(
            "chart.set-data-table",
            ServiceRegistry.Chart.RouteCliArgs("set-data-table", slideIndex: 1, shapeIndex: 1, visible: true).Command);
    }

    [Fact]
    public void ChartQuickFormattingResult_SerializesPairedReadWriteFields()
    {
        var result = new ChartOperationResult
        {
            Success = true,
            ShapeIndex = 1,
            ChartStyle = 2,
            ColorStyle = 3,
            HasDataTable = true
        };

        using var document = JsonDocument.Parse(PowerPointToolsBase.Serialize(result));
        JsonElement root = document.RootElement;

        Assert.Equal(2, root.GetProperty("chartStyle").GetInt32());
        Assert.Equal(3, root.GetProperty("colorStyle").GetInt32());
        Assert.True(root.GetProperty("hasDataTable").GetBoolean());
    }

    [Fact]
    public void ShapeLinkActions_HaveGeneratedCliAndServiceWiring()
    {
        var expectedActions = new[]
        {
            "get-link-info",
            "update-link",
            "break-link",
            "set-link-auto-update"
        };

        foreach (string action in expectedActions)
        {
            Assert.Contains(action, ServiceRegistry.Shape.ValidActions);
        }

        Assert.Equal(
            "shape.get-link-info",
            ServiceRegistry.Shape.RouteCliArgs(
                "get-link-info", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "shape.update-link",
            ServiceRegistry.Shape.RouteCliArgs(
                "update-link", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "shape.break-link",
            ServiceRegistry.Shape.RouteCliArgs(
                "break-link", slideIndex: 1, shapeIndex: 1).Command);
        Assert.Equal(
            "shape.set-link-auto-update",
            ServiceRegistry.Shape.RouteCliArgs(
                "set-link-auto-update", slideIndex: 1, shapeIndex: 1, autoUpdate: true).Command);
    }

    [Fact]
    public void ShapeCli_RoutesWordArtAnd3DRotationActions()
    {
        string[] expectedActions = ["add-text-effect", "set-3d-rotation", "get-3d-rotation"];

        Assert.All(expectedActions, action => Assert.Contains(action, ServiceRegistry.Shape.ValidActions));

        Assert.Equal(
            "shape.add-text-effect",
            ServiceRegistry.Shape.RouteCliArgs(
                "add-text-effect",
                slideIndex: 1,
                presetEffect: "msoTextEffect1",
                text: "Quarterly outlook",
                fontName: "Arial",
                fontSize: 36f,
                left: 40f,
                top: 50f).Command);

        Assert.Equal(
            "shape.set-3d-rotation",
            ServiceRegistry.Shape.RouteCliArgs(
                "set-3d-rotation",
                slideIndex: 1,
                shapeIndex: 1,
                rotationX: 20f,
                rotationY: -30f,
                rotationZ: 40f).Command);

        Assert.Equal(
            "shape.get-3d-rotation",
            ServiceRegistry.Shape.RouteCliArgs(
                "get-3d-rotation", slideIndex: 1, shapeIndex: 1).Command);
    }

    [Fact]
    public void Shape3DRotationCli_AcceptsIndividualAxesAndRejectsInapplicableOnes()
    {
        foreach (var axis in new (string Name, float? X, float? Y, float? Z)[]
        {
            ("rotationX", 20f, null, null),
            ("rotationY", null, -30f, null),
            ("rotationZ", null, null, 40f)
        })
        {
            var (command, args) = ServiceRegistry.Shape.RouteCliArgs(
                "set-3d-rotation",
                slideIndex: 1,
                shapeIndex: 1,
                rotationX: axis.X,
                rotationY: axis.Y,
                rotationZ: axis.Z);

            Assert.Equal("shape.set-3d-rotation", command);
            Assert.Contains($"\"{axis.Name}\":", JsonSerializer.Serialize(args), StringComparison.Ordinal);
        }

        // The read action takes no axis arguments, and the 2D action must not accept 3D ones.
        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.RouteCliArgs(
                "get-3d-rotation", slideIndex: 1, shapeIndex: 1, rotationX: 20f));

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.RouteCliArgs(
                "set-rotation", slideIndex: 1, shapeIndex: 1, degrees: 15f, rotationZ: 40f));
    }

    [Fact]
    public void AddAttachedConnector_HasGeneratedCliWiringAndRequiredParameters()
    {
        Assert.Contains("add-attached-connector", ServiceRegistry.Shape.ValidActions);

        Assert.Equal(
            "shape.add-attached-connector",
            ServiceRegistry.Shape.RouteCliArgs(
                "add-attached-connector",
                slideIndex: 1,
                connectorType: "msoConnectorStraight",
                beginShapeIndex: 1,
                beginConnectionSite: 2,
                endShapeIndex: 2,
                endConnectionSite: 4).Command);

        var error = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.ValidateActionArguments(
                "add-attached-connector",
                """{"slideIndex":1,"connectorType":"msoConnectorStraight","beginShapeIndex":1,"beginConnectionSite":2,"endShapeIndex":2}"""));

        Assert.Contains("endConnectionSite", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ShapeCopyFormatting_HasGeneratedCliAndServiceWiring()
    {
        Assert.Contains("copy-formatting", ServiceRegistry.Shape.ValidActions);

        Assert.Equal(
            "shape.copy-formatting",
            ServiceRegistry.Shape.RouteCliArgs(
                "copy-formatting",
                slideIndex: 1,
                sourceShapeIndex: 2,
                targetShapeIndex: 3).Command);

        ServiceRegistry.Shape.ValidateActionArguments(
            "copy-formatting",
            """{"slideIndex":1,"sourceShapeIndex":2,"targetShapeIndex":3}""");
    }

    [Fact]
    public void ShapeMerge_HasGeneratedCliAndServiceWiring()
    {
        Assert.Contains("merge", ServiceRegistry.Shape.ValidActions);

        Assert.Equal(
            "shape.merge",
            ServiceRegistry.Shape.RouteCliArgs(
                "merge",
                slideIndex: 1,
                shapeIndexes: [1, 2],
                mergeType: "msoMergeUnion").Command);

        ServiceRegistry.Shape.ValidateActionArguments(
            "merge",
            """{"slideIndex":1,"shapeIndexes":[1,2],"mergeType":"msoMergeUnion"}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.ValidateActionArguments(
                "merge",
                """{"slideIndex":1,"shapeIndexes":[1,2]}"""));

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.ValidateActionArguments(
                "copy-formatting",
                """{"slideIndex":1,"sourceShapeIndex":2}"""));
    }

    [Fact]
    public void ShapeDuplicate_HasGeneratedCliAndServiceWiring()
    {
        Assert.Contains("duplicate", ServiceRegistry.Shape.ValidActions);

        Assert.Equal(
            "shape.duplicate",
            ServiceRegistry.Shape.RouteCliArgs(
                "duplicate",
                slideIndex: 1,
                shapeIndex: 2).Command);

        ServiceRegistry.Shape.ValidateActionArguments(
            "duplicate",
            """{"slideIndex":1,"shapeIndex":2}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.ValidateActionArguments(
                "duplicate",
                """{"slideIndex":1}"""));
    }

    [Fact]
    public void ShapeCopyToSlide_HasGeneratedCliAndServiceWiring()
    {
        Assert.Contains("copy-to-slide", ServiceRegistry.Shape.ValidActions);

        Assert.Equal(
            "shape.copy-to-slide",
            ServiceRegistry.Shape.RouteCliArgs(
                "copy-to-slide",
                slideIndex: 1,
                shapeIndex: 2,
                targetSlideIndex: 3).Command);

        ServiceRegistry.Shape.ValidateActionArguments(
            "copy-to-slide",
            """{"slideIndex":1,"shapeIndex":2,"targetSlideIndex":3}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Shape.ValidateActionArguments(
                "copy-to-slide",
                """{"slideIndex":1,"shapeIndex":2}"""));
    }

    [Fact]
    public void CustomShow_HasGeneratedCliAndServiceWiring()
    {
        Assert.Contains("list", ServiceRegistry.CustomShow.ValidActions);
        Assert.Contains("create", ServiceRegistry.CustomShow.ValidActions);
        Assert.Contains("delete", ServiceRegistry.CustomShow.ValidActions);

        Assert.Equal(
            "customshow.create",
            ServiceRegistry.CustomShow.RouteCliArgs(
                "create",
                name: "Demo",
                slideIndices: [1, 3]).Command);

        ServiceRegistry.CustomShow.ValidateActionArguments(
            "create",
            """{"name":"Demo","slideIndices":[1,3]}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.CustomShow.ValidateActionArguments(
                "create",
                """{"name":"Demo"}"""));

        ServiceRegistry.CustomShow.ValidateActionArguments("delete", """{"name":"Demo"}""");

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.CustomShow.ValidateActionArguments("delete", """{}"""));
    }
}
