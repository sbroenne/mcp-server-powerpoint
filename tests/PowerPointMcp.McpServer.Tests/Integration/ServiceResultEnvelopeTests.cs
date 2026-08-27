using System.Text.Json;
using Sbroenne.PowerPointMcp.McpServer.Infrastructure;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Verifies generated Core result envelopes are reflected in the shared service response without
/// touching PowerPoint COM.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Service")]
[Trait("Feature", "McpProtocol")]
public sealed class ServiceResultEnvelopeTests
{
    [Fact]
    public void WrapResult_WithCoreValidationFailure_PropagatesFailure()
    {
        const string result =
            """{"success":false,"errorMessage":"Validation failed.","itemIndex":7}""";

        ServiceResponse response = PowerPointMcpService.WrapResult(result);

        Assert.False(response.Success);
        Assert.Equal("Validation failed.", response.ErrorMessage);
        Assert.Equal(result, response.Result);
    }

    [Fact]
    public void WrapResult_WithSuccessfulCoreResult_PreservesPayload()
    {
        const string result = """{"success":true,"shapeIndex":1}""";

        ServiceResponse response = PowerPointMcpService.WrapResult(result);

        Assert.True(response.Success);
        Assert.Null(response.ErrorMessage);
        Assert.Equal(result, response.Result);
    }

    [Fact]
    public void ServiceBridge_WithCoreValidationFailure_KeepsStableMcpErrorEnvelope()
    {
        const string result =
            """{"success":false,"errorMessage":"Validation failed.","itemIndex":7}""";
        var response = new ServiceResponse
        {
            Success = false,
            ErrorMessage = "Validation failed.",
            Result = result
        };

        string json = ServiceBridge.FormatResponse(response);

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("Validation failed.", root.GetProperty("errorMessage").GetString());
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.False(root.TryGetProperty("itemIndex", out _));
    }
}
