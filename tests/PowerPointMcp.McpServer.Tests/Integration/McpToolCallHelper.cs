// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Shared JSON tools/call helpers for tests that drive the MCP server exclusively through the
/// official <see cref="McpClient"/> — asserting responses only via the JSON payload, never via
/// direct method calls. Used by <see cref="McpAuthoringWorkflowTests"/> and
/// <see cref="McpShutdownRobustnessTests"/> (kept out of <see cref="McpRoundTripTests"/> so that
/// file's existing, already-passing helpers are left untouched).
/// </summary>
internal static class McpToolCallHelper
{
    /// <summary>Calls a tool via MCP protocol and returns the raw JSON text response.</summary>
    public static async Task<string> CallToolAsync(
        McpClient client,
        string toolName,
        Dictionary<string, object?> arguments,
        CancellationToken cancellationToken)
    {
        var result = await client.CallToolAsync(toolName, arguments, cancellationToken: cancellationToken);

        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        return textBlock.Text;
    }

    /// <summary>Fails the test with a clear message if the tool response has success=false.</summary>
    public static void AssertSuccess(string jsonResult, string operationName)
    {
        JsonDocument json;
        try
        {
            json = JsonDocument.Parse(jsonResult);
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
            return;
        }

        using (json)
        {
            if (json.RootElement.TryGetProperty("success", out var success) && !success.GetBoolean())
            {
                var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                    ? errProp.GetString()
                    : "Unknown error";
                Assert.Fail($"{operationName} returned success=false: {errorMsg}\nResponse: {jsonResult}");
            }
        }
    }

    public static string? GetString(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }

    public static int? GetInt(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) && prop.ValueKind != JsonValueKind.Null
            ? prop.GetInt32()
            : null;
    }

    public static double? GetDouble(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) && prop.ValueKind != JsonValueKind.Null
            ? prop.GetDouble()
            : null;
    }

    public static bool? GetBool(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) && prop.ValueKind != JsonValueKind.Null
            ? prop.GetBoolean()
            : null;
    }
}
