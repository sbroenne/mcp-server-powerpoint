namespace Sbroenne.PowerPointMcp.Core.Tags;

/// <summary>A PowerPoint string tag in its native 1-based collection order.</summary>
public sealed class TagInfo
{
    /// <summary>1-based index in the owner's native tag collection.</summary>
    public required int TagIndex { get; init; }

    /// <summary>Normalized, case-insensitive tag name.</summary>
    public required string Name { get; init; }

    /// <summary>Tag value exactly as stored in PowerPoint.</summary>
    public required string Value { get; init; }
}
