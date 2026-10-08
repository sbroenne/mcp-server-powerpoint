namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>Bounded metadata and text preview for one slide.</summary>
public sealed class SlideOverviewInfo
{
    /// <summary>The slide's 1-based position in the presentation.</summary>
    public int SlideIndex { get; init; }

    /// <summary>The slide's PowerPoint name.</summary>
    public string? Name { get; init; }

    /// <summary>The slide's custom layout name.</summary>
    public string? LayoutName { get; init; }

    /// <summary>Number of top-level shapes on the slide.</summary>
    public int ShapeCount { get; init; }

    /// <summary>Text from the slide's shapes, limited by the request's per-slide character cap.</summary>
    public string? TextPreview { get; init; }

    /// <summary>Whether the preview was shortened by its character or shape-scan limit.</summary>
    public bool TextTruncated { get; init; }
}
