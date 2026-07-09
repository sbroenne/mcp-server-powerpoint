namespace Sbroenne.PowerPointMcp.Core.Attributes;

/// <summary>
/// Specifies which MCP tool exposes this interface or method.
/// Used by code generator to group methods into MCP tools and generate MCP tool classes.
/// </summary>
/// <remarks>
/// Can be applied at interface level (all methods go to same tool)
/// or method level (methods can be split across different tools).
/// Method-level attribute overrides interface-level.
/// </remarks>
[AttributeUsage(AttributeTargets.Interface | AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class McpToolAttribute : Attribute
{
    /// <summary>
    /// The MCP tool name (e.g., "slide", "shape").
    /// </summary>
    public string ToolName { get; }

    /// <summary>
    /// Human-readable title for the MCP tool (e.g., "Slide Operations").
    /// Used in [McpServerTool(Title = ...)].
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Whether the tool is destructive (modifies data). Default: true.
    /// Used in [McpServerTool(Destructive = ...)].
    /// </summary>
    public bool Destructive { get; set; } = true;

    /// <summary>
    /// MCP meta category for the tool (e.g., "data", "analysis", "query", "settings").
    /// Used in [McpMeta("category", ...)].
    /// </summary>
    public string? Category { get; set; }

    /// <summary>
    /// Tool description shown to LLMs via [Description("...")].
    /// Since source generators can't read XML docs from metadata references,
    /// this provides the description that appears in the MCP JSON schema.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// When true, <c>PowerPointMcp.Generators.Mcp</c> skips generating an action-dispatch MCP
    /// tool for this category, even though its CLI/service registry code is still generated
    /// normally. Use this for categories whose MCP surface is intentionally hand-written instead
    /// (e.g. session lifecycle tools for Presentation — <c>create_presentation</c>,
    /// <c>open_presentation</c>, etc. — which have no CLI-parity concept of a generated
    /// "presentation" action-dispatch tool). Default: false.
    /// </summary>
    public bool SkipMcpToolGeneration { get; set; }

    /// <summary>
    /// Creates a new McpToolAttribute.
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "slide")</param>
    public McpToolAttribute(string toolName)
    {
        ToolName = toolName ?? throw new ArgumentNullException(nameof(toolName));
    }
}
