using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.PageSetup;

/// <summary>
/// Presentation-wide slide size, numbering, and footer settings.
/// </summary>
[ServiceCategory("pagesetup", "PageSetup")]
[McpTool("pagesetup", Title = "Page Setup Operations", Destructive = true, Category = "content",
    Description = "Read or change presentation-wide slide size, numbering, and footer settings.")]
public interface IPageSetupCommands
{
    /// <summary>Gets slide dimensions, orientation, and the first slide number.</summary>
    PageSetupOperationResult GetSettings(IPresentationBatch batch);

    /// <summary>Sets custom slide dimensions in points.</summary>
    PageSetupOperationResult SetSize(IPresentationBatch batch, float width, float height);

    /// <summary>Sets the number displayed on the first slide.</summary>
    PageSetupOperationResult SetFirstSlideNumber(IPresentationBatch batch, int firstSlideNumber);

    /// <summary>Gets presentation-wide footer, slide-number, and date/time settings.</summary>
    PageSetupOperationResult GetFooter(IPresentationBatch batch);

    /// <summary>
    /// Updates presentation-wide footer settings. Null values leave the corresponding setting
    /// unchanged. <paramref name="dateTimeMode"/> is <c>automatic</c> or <c>fixed</c>.
    /// </summary>
    PageSetupOperationResult SetFooter(
        IPresentationBatch batch,
        string? footerText = null,
        bool? showFooter = null,
        bool? showSlideNumber = null,
        bool? showDateTime = null,
        string? dateTimeMode = null,
        string? fixedDateTimeText = null,
        bool? showOnTitleSlide = null);
}
