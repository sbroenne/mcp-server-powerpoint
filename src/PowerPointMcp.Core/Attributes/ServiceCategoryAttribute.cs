namespace Sbroenne.PowerPointMcp.Core.Attributes;

/// <summary>
/// Marks an interface as a service category for code generation.
/// The category name is used for Service routing (e.g., "slide" → "slide.add").
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class ServiceCategoryAttribute : Attribute
{
    /// <summary>
    /// The service category name (e.g., "slide", "shape", "table").
    /// Used in service command routing: "{category}.{action}"
    /// </summary>
    public string Category { get; }

    /// <summary>
    /// Optional PascalCase name for the category (e.g., "TextFrame").
    /// If not specified, derived from category using simple conversion.
    /// Use this when the category contains multiple words (e.g., "textframe" → "TextFrame").
    /// </summary>
    public string? PascalName { get; }

    /// <summary>
    /// Creates a new ServiceCategoryAttribute.
    /// </summary>
    /// <param name="category">The service category name in lowercase (e.g., "slide")</param>
    /// <param name="pascalName">Optional PascalCase name (e.g., "TextFrame")</param>
    public ServiceCategoryAttribute(string category, string? pascalName = null)
    {
        Category = category ?? throw new ArgumentNullException(nameof(category));
        PascalName = pascalName;
    }
}
