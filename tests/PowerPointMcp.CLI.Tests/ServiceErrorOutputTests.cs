using System.Text.Json;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class ServiceErrorOutputTests
{
    [Fact]
    public void SerializeServiceError_WithStructuredCoreFailure_KeepsStableCliEnvelope()
    {
        var response = new ServiceResponse
        {
            Success = false,
            ErrorMessage = "Validation failed.",
            Result = """{"success":false,"errorMessage":"Validation failed.","itemIndex":7}"""
        };

        string json = CliErrorOutput.SerializeServiceError(response);

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("Validation failed.", root.GetProperty("errorMessage").GetString());
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.False(root.TryGetProperty("itemIndex", out _));
    }
}
