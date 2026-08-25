using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Accessibility;

/// <summary>Deterministic presentation accessibility checks and reading-order operations.</summary>
[ServiceCategory("accessibility", "Accessibility")]
[McpTool("accessibility", Title = "Accessibility Operations", Destructive = true, Category = "content",
    Description = "Audit presentation accessibility and read or change a slide's shape reading order.")]
public interface IAccessibilityCommands
{
    /// <summary>Audits the presentation for deterministic accessibility issues.</summary>
    AccessibilityOperationResult Audit(IPresentationBatch batch);

    /// <summary>Gets a slide's shape reading order as 1-based shape indexes.</summary>
    AccessibilityOperationResult GetReadingOrder(IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Sets a slide's reading order. <paramref name="shapeIndexes"/> must contain every 1-based
    /// shape index on the slide exactly once.
    /// </summary>
    AccessibilityOperationResult SetReadingOrder(
        IPresentationBatch batch,
        int slideIndex,
        IReadOnlyList<int> shapeIndexes);
}
