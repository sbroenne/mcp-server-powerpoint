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
        const string result = """{"success":false,"errorMessage":"Invalid media path."}""";

        ServiceResponse response = PowerPointMcpService.WrapResult(result);

        Assert.False(response.Success);
        Assert.Equal("Invalid media path.", response.ErrorMessage);
        Assert.Null(response.Result);
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
}
