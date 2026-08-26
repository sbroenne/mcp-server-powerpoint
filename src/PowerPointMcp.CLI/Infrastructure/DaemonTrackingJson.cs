using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

internal static class DaemonTrackingJson
{
    internal static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        WriteIndented = true,
    };
}
