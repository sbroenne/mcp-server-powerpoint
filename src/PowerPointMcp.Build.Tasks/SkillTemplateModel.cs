namespace Sbroenne.PowerPointMcp.Build.Tasks;

/// <summary>
/// Model passed to Scriban templates for skill generation.
/// Properties are exposed to templates using snake_case naming (Scriban convention).
/// </summary>
public class SkillTemplateModel
{
    /// <summary>Number of CLI commands or MCP tools</summary>
    public int ToolCount { get; set; }

    /// <summary>Total number of operations/actions</summary>
    public int OperationCount { get; set; }

    /// <summary>For CLI skill: parsed command reference from the generator's manifest</summary>
    public List<CliCommand>? CliCommands { get; set; }
}

/// <summary>
/// Represents a CLI command group generated from a Core <c>[ServiceCategory]</c> interface
/// (e.g. "chart", "shape", "table").
/// </summary>
public class CliCommand
{
    /// <summary>Command name (e.g., "shape", "table")</summary>
    public string Name { get; set; } = "";

    /// <summary>Command description from interface XML doc</summary>
    public string Description { get; set; } = "";

    /// <summary>List of actions (e.g., "add-rectangle", "delete")</summary>
    public List<string> Actions { get; set; } = new();

    /// <summary>List of parameters</summary>
    public List<CliParameter> Parameters { get; set; } = new();
}

/// <summary>
/// Represents a CLI parameter for a command group.
/// </summary>
public class CliParameter
{
    /// <summary>Parameter name without dashes (e.g., "slide-index")</summary>
    public string Name { get; set; } = "";

    /// <summary>Description, including which action(s) require it</summary>
    public string Description { get; set; } = "";

    /// <summary>Whether parameter has short form (e.g., -s for --session) — unused for pptcli today</summary>
    public string? ShortForm { get; set; }
}
