using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Master;

/// <summary>
/// Slide master commands: read/edit the title and body placeholder fonts on the presentation's
/// slide master, and read/edit the slide master's background fill color. Operates within an
/// already-open <see cref="IPresentationBatch"/>. Changes here apply to every slide that
/// inherits from the master (i.e. any slide that does not itself override the property), which
/// is the practical "edit the master, not each slide" workflow PowerPoint's COM object model
/// supports safely.
/// </summary>
/// <remarks>
/// This is a narrower surface than a full slide-master editor (no custom-layout authoring, no
/// multi-master/theme-swap here — that is <see cref="Presentation.IPresentationCommands.ApplyTemplate"/>'s
/// job). It targets the two things authors most commonly want to change on a master: the
/// title/body placeholder font, and the master background color.
/// </remarks>
[ServiceCategory("master", "Master")]
[McpTool("master", Title = "Slide Master Operations", Destructive = true, Category = "content",
    Description = "Read or edit the slide master's title/body placeholder fonts and background color in an open presentation session. Changes apply to every slide inheriting from the master.")]
public interface IMasterCommands
{
    /// <summary>Gets the font name, size, bold, and color of the master's title placeholder.</summary>
    MasterOperationResult GetTitleFont(IPresentationBatch batch);

    /// <summary>
    /// Sets one or more font properties of the master's title placeholder. Any parameter left
    /// null is unchanged.
    /// </summary>
    MasterOperationResult SetTitleFont(
        IPresentationBatch batch,
        string? fontName = null,
        float? fontSize = null,
        bool? bold = null,
        byte? red = null,
        byte? green = null,
        byte? blue = null);

    /// <summary>Gets the font name, size, bold, and color of the master's body placeholder.</summary>
    MasterOperationResult GetBodyFont(IPresentationBatch batch);

    /// <summary>
    /// Sets one or more font properties of the master's body placeholder. Any parameter left
    /// null is unchanged.
    /// </summary>
    MasterOperationResult SetBodyFont(
        IPresentationBatch batch,
        string? fontName = null,
        float? fontSize = null,
        bool? bold = null,
        byte? red = null,
        byte? green = null,
        byte? blue = null);

    /// <summary>Gets the slide master's background fill color.</summary>
    MasterOperationResult GetBackgroundColor(IPresentationBatch batch);

    /// <summary>Sets the slide master's background fill to a solid RGB color.</summary>
    MasterOperationResult SetBackgroundColor(IPresentationBatch batch, byte red, byte green, byte blue);

    /// <summary>
    /// Sets a two-color gradient background for the slide master. <paramref name="gradientStyle"/>
    /// is an <c>MsoGradientStyle</c> member name (e.g. "msoGradientHorizontal", "msoGradientVertical",
    /// "msoGradientDiagonalUp", "msoGradientDiagonalDown", "msoGradientFromCorner",
    /// "msoGradientFromTitle", "msoGradientFromCenter"; defaults to "msoGradientHorizontal").
    /// <paramref name="gradientVariant"/> selects one of PowerPoint's 1-4 preset variants for
    /// that style (defaults to 1).
    /// </summary>
    MasterOperationResult SetGradientBackground(
        IPresentationBatch batch,
        byte red1, byte green1, byte blue1,
        byte red2, byte green2, byte blue2,
        string gradientStyle = "msoGradientHorizontal",
        int gradientVariant = 1);

    /// <summary>
    /// Gets the slide master's gradient background: both stop colors, the <c>MsoGradientStyle</c>
    /// member name, and the variant. Fails if the master's background is not currently a
    /// gradient fill.
    /// </summary>
    MasterOperationResult GetGradientBackground(IPresentationBatch batch);
}
