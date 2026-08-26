namespace Sbroenne.PowerPointMcp.Core.Tags;

internal sealed class TagCollectionResult
{
    public bool Success { get; init; }

    public string? ErrorMessage { get; init; }

    public string? Name { get; init; }

    public string? Value { get; init; }

    public int? Index { get; init; }

    public int Count { get; init; }

    public IReadOnlyList<TagInfo>? Tags { get; init; }
}
