namespace Sbroenne.PowerPointMcp.Core.PageSetup;

/// <summary>Result of a presentation page-setup operation.</summary>
public sealed class PageSetupOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Slide width in points.</summary>
    public float? Width { get; init; }

    /// <summary>Slide height in points.</summary>
    public float? Height { get; init; }

    /// <summary>Derived orientation: <c>landscape</c>, <c>portrait</c>, or <c>square</c>.</summary>
    public string? Orientation { get; init; }

    /// <summary>The number displayed on the first slide.</summary>
    public int? FirstSlideNumber { get; init; }

    /// <summary>Footer text.</summary>
    public string? FooterText { get; init; }

    /// <summary>Whether the footer is visible.</summary>
    public bool? ShowFooter { get; init; }

    /// <summary>Whether slide numbers are visible.</summary>
    public bool? ShowSlideNumber { get; init; }

    /// <summary>Whether date/time is visible.</summary>
    public bool? ShowDateTime { get; init; }

    /// <summary>Date/time mode: <c>automatic</c> or <c>fixed</c>.</summary>
    public string? DateTimeMode { get; init; }

    /// <summary>Fixed date/time text when fixed mode is selected.</summary>
    public string? FixedDateTimeText { get; init; }

    /// <summary>Whether footer elements are shown on title slides.</summary>
    public bool? ShowOnTitleSlide { get; init; }
}
