using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

internal sealed class PresentationToolOutputSchema
{
    public bool Success { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SessionId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? PresentationPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Count { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Closed { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PresentationSessionOutputSchema>? Sessions { get; set; }
}

internal sealed class PresentationSessionOutputSchema
{
    public string SessionId { get; set; } = string.Empty;
    public string PresentationPath { get; set; } = string.Empty;
    public bool IsPowerPointProcessAlive { get; set; }
}
