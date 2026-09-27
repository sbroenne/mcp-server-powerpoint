namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>A literal match in the original, unmodified shape text frame.</summary>
public sealed class TextMatch
{
    /// <summary>1-based PowerPoint character position within the text frame.</summary>
    public int Start { get; init; }

    /// <summary>Length in PowerPoint text-range characters.</summary>
    public int Length { get; init; }

    /// <summary>The matched text with its original casing.</summary>
    public string Text { get; init; } = string.Empty;
}