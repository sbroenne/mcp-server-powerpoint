namespace Sbroenne.PowerPointMcp.Core.Layout;

/// <summary>
/// Result of a slide layout operation (set/get layout).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1).
/// </remarks>
public sealed class LayoutOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// The slide layout name (a <c>Microsoft.Office.Interop.PowerPoint.PpSlideLayout</c>
    /// enum member name, e.g. "ppLayoutBlank", "ppLayoutTitleOnly", "ppLayoutText").
    /// </summary>
    public string? LayoutName { get; init; }
}
