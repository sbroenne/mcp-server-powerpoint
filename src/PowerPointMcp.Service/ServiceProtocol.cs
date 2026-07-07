using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Service;

/// <summary>
/// Protocol messages for CLI-to-daemon communication over named pipes (StreamJsonRpc transport,
/// JSON-RPC 2.0). Ported from mcp-server-excel's ExcelMcp.Service.ServiceProtocol.
/// </summary>
public static class ServiceProtocol
{
    /// <summary>Shared JSON options for request/response/result payloads (camelCase, omit nulls, string enums).</summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };
}

/// <summary>Request sent from the CLI to the daemon.</summary>
public sealed class ServiceRequest
{
    /// <summary>Command to execute (e.g. "session.open", "notes.set-notes-text").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on an open presentation.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }

    /// <summary>Source of the request (e.g. "cli").</summary>
    public string? Source { get; init; }
}

/// <summary>Response sent from the daemon back to the CLI.</summary>
public sealed class ServiceResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>The command that produced this response, when available.</summary>
    public string? Command { get; init; }

    /// <summary>The session ID associated with this response, when available.</summary>
    public string? SessionId { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Structured error category if Success is false.</summary>
    public string? ErrorCategory { get; init; }

    /// <summary>Exception type that produced the failure, when available.</summary>
    public string? ExceptionType { get; init; }

    /// <summary>HRESULT from a COM failure, when available.</summary>
    [JsonPropertyName("hresult")]
    public string? HResult { get; init; }

    /// <summary>Inner exception details, when available.</summary>
    public string? InnerError { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}

/// <summary>Daemon status information, returned by "service.status".</summary>
public sealed class ServiceStatus
{
    /// <summary>Whether the daemon is currently running (always true for a live response).</summary>
    public bool Running { get; init; }

    /// <summary>The daemon process's OS process id.</summary>
    public int ProcessId { get; init; }

    /// <summary>Number of currently open presentation sessions.</summary>
    public int SessionCount { get; init; }

    /// <summary>UTC time the daemon started.</summary>
    public DateTime StartTime { get; init; }

    /// <summary>How long the daemon has been running.</summary>
    public TimeSpan Uptime => Running ? DateTime.UtcNow - StartTime : TimeSpan.Zero;
}
