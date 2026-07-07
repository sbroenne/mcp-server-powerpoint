using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Writes structured JSON error envelopes to stdout for CLI failures. Ported from
/// mcp-server-excel's ExcelMcp.CLI.Infrastructure.CliErrorOutput.
/// </summary>
internal static class CliErrorOutput
{
    /// <summary>Writes an error envelope for an unhandled local exception (not from the daemon).</summary>
    public static int WriteException(Exception ex, string? errorCategory = null)
    {
        Console.WriteLine(Serialize(
            ex.Message,
            errorCategory,
            null,
            null,
            ex.GetType().Name,
            null,
            ex.InnerException?.Message));
        return 1;
    }

    /// <summary>Writes an error envelope from a failed <see cref="ServiceResponse"/>.</summary>
    public static int WriteServiceError(ServiceResponse response)
    {
        Console.WriteLine(Serialize(
            response.ErrorMessage,
            response.ErrorCategory,
            response.Command,
            response.SessionId,
            response.ExceptionType,
            response.HResult,
            response.InnerError));
        return 1;
    }

    /// <summary>Writes a plain error envelope with just a message.</summary>
    public static int WriteError(string errorMessage, string? errorCategory = null)
    {
        Console.WriteLine(Serialize(errorMessage, errorCategory, null, null, null, null, null));
        return 1;
    }

    private static string Serialize(
        string? errorMessage,
        string? errorCategory,
        string? command,
        string? sessionId,
        string? exceptionType,
        string? hresult,
        string? innerError)
    {
        return JsonSerializer.Serialize(new ErrorEnvelope
        {
            Success = false,
            Error = errorMessage ?? "Unknown error.",
            ErrorMessage = errorMessage ?? "Unknown error.",
            ErrorCategory = errorCategory,
            Command = command,
            SessionId = sessionId,
            IsError = true,
            ExceptionType = exceptionType,
            HResult = hresult,
            InnerError = innerError
        }, ServiceProtocol.JsonOptions);
    }

    private sealed class ErrorEnvelope
    {
        public bool Success { get; init; }

        public string Error { get; init; } = string.Empty;

        public string ErrorMessage { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ErrorCategory { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Command { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SessionId { get; init; }

        public bool IsError { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ExceptionType { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("hresult")]
        public string? HResult { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? InnerError { get; init; }
    }
}
