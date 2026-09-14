namespace Sbroenne.PowerPointMcp.Core.Attributes;

/// <summary>
/// Overrides the action name generated from a service method name.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class ServiceActionAttribute : Attribute
{
    /// <summary>Creates an action-name override.</summary>
    public ServiceActionAttribute(string action)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(action);
        Action = action;
    }

    /// <summary>The kebab-case action name exposed by generated entry points.</summary>
    public string Action { get; }
}