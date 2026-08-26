using Sbroenne.PowerPointMcp.Generated;

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
}
