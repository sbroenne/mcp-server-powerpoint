namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>One legacy PowerPoint slide comment.</summary>
public sealed class SlideCommentInfo
{
    /// <summary>1-based index within the slide's legacy comment collection.</summary>
    public required int CommentIndex { get; init; }

    /// <summary>Comment author.</summary>
    public required string Author { get; init; }

    /// <summary>Comment author initials.</summary>
    public required string Initials { get; init; }

    /// <summary>Comment text.</summary>
    public required string Text { get; init; }

    /// <summary>Horizontal comment marker position in points.</summary>
    public required float Left { get; init; }

    /// <summary>Vertical comment marker position in points.</summary>
    public required float Top { get; init; }
}
