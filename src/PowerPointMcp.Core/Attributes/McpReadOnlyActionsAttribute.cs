namespace Sbroenne.PowerPointMcp.Core.Attributes;

/// <summary>
/// Exposes the named service actions through a separate read-only MCP tool.
/// The generated tool name appends "_read" to the interface's MCP tool name.
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class McpReadOnlyActionsAttribute(params string[] actionNames) : Attribute
{
    internal IReadOnlyList<string> ActionNames { get; } = actionNames;
}
