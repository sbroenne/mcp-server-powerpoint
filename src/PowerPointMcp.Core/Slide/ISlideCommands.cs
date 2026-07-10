using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>
/// Slide lifecycle commands: add, delete, count, duplicate, reorder, per-slide background color,
/// and section management. First domain built on top of the presentation lifecycle commands,
/// operating within an already-open IPresentationBatch.
/// </summary>
[ServiceCategory("slide", "Slide")]
[McpTool("slide", Title = "Slide Operations", Destructive = true, Category = "content",
    Description = "Add, count, delete, duplicate, reorder, and set section/background on slides in an open presentation session.")]
public interface ISlideCommands
{
    /// <summary>
    /// Adds a new blank slide at the end of the presentation.
    /// </summary>
    /// <returns>A result with Success, the new slide's 1-based index, and the new total count.</returns>
    SlideOperationResult AddBlank(IPresentationBatch batch);

    /// <summary>
    /// Gets the current number of slides in the presentation.
    /// </summary>
    SlideOperationResult GetCount(IPresentationBatch batch);

    /// <summary>
    /// Deletes the slide at the given 1-based index.
    /// </summary>
    SlideOperationResult Delete(IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Duplicates the slide at the given 1-based index. The duplicate is inserted immediately
    /// after the source slide. Returns the duplicate's new 1-based <c>slideIndex</c>.
    /// </summary>
    SlideOperationResult Duplicate(IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Moves the slide at <paramref name="slideIndex"/> to the 1-based position
    /// <paramref name="toPosition"/>, renumbering all other slides accordingly.
    /// </summary>
    SlideOperationResult MoveTo(IPresentationBatch batch, int slideIndex, int toPosition);

    /// <summary>
    /// Sets a solid background color for a single slide, overriding the slide master's
    /// background for that slide only (sets <c>FollowMasterBackground = False</c>).
    /// </summary>
    SlideOperationResult SetBackgroundColor(IPresentationBatch batch, int slideIndex, byte red, byte green, byte blue);

    /// <summary>
    /// Gets a slide's background color and whether it currently follows the slide master's
    /// background (<c>followsMasterBackground</c>).
    /// </summary>
    SlideOperationResult GetBackgroundColor(IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Sets a two-color gradient background for a single slide, overriding the slide master's
    /// background for that slide only (sets <c>FollowMasterBackground = False</c>).
    /// <paramref name="gradientStyle"/> is an <c>MsoGradientStyle</c> member name (e.g.
    /// "msoGradientHorizontal", "msoGradientVertical", "msoGradientDiagonalUp",
    /// "msoGradientDiagonalDown", "msoGradientFromCorner", "msoGradientFromTitle",
    /// "msoGradientFromCenter"; defaults to "msoGradientHorizontal"). <paramref name="gradientVariant"/>
    /// selects one of PowerPoint's 1-4 preset variants for that style (defaults to 1).
    /// </summary>
    SlideOperationResult SetGradientBackground(
        IPresentationBatch batch,
        int slideIndex,
        byte red1, byte green1, byte blue1,
        byte red2, byte green2, byte blue2,
        string gradientStyle = "msoGradientHorizontal",
        int gradientVariant = 1);

    /// <summary>
    /// Gets a slide's gradient background: both stop colors, the <c>MsoGradientStyle</c> member
    /// name, and the variant. Fails if the slide's background is not currently a gradient fill.
    /// </summary>
    SlideOperationResult GetGradientBackground(IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Adds a new section before the given 1-based <paramref name="sectionIndex"/>. Pass
    /// <c>sectionCount + 1</c> to append a section at the end. Returns the new section's
    /// <c>sectionIndex</c> and the new total <c>sectionCount</c>.
    /// </summary>
    SlideOperationResult AddSection(IPresentationBatch batch, int sectionIndex, string? sectionName = null);

    /// <summary>Renames the section at the given 1-based <paramref name="sectionIndex"/>.</summary>
    SlideOperationResult RenameSection(IPresentationBatch batch, int sectionIndex, string sectionName);

    /// <summary>
    /// Deletes the section at the given 1-based <paramref name="sectionIndex"/>. If
    /// <paramref name="deleteSlides"/> is true, the slides in that section are deleted too;
    /// otherwise they are kept and become part of the neighboring section. Note: PowerPoint
    /// disallows deleting section 1 unless <paramref name="deleteSlides"/> is true.
    /// </summary>
    SlideOperationResult DeleteSection(IPresentationBatch batch, int sectionIndex, bool deleteSlides = false);

    /// <summary>Gets the current number of sections in the presentation.</summary>
    SlideOperationResult GetSectionCount(IPresentationBatch batch);

    /// <summary>Gets the name of the section at the given 1-based <paramref name="sectionIndex"/>.</summary>
    SlideOperationResult GetSectionName(IPresentationBatch batch, int sectionIndex);
}
