using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Chart;
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
}
