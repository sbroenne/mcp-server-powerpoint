using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Animation;

/// <summary>
/// Animation commands: add/delete entrance, emphasis, and exit effects on a shape's slide
/// timeline (<c>Slide.TimeLine.MainSequence</c>), and read/set a slide's transition
/// (<c>Slide.SlideShowTransition</c>). Operates within an already-open
/// <see cref="ComInterop.Session.IPresentationBatch"/>, targeting a specific slide (and, for
/// shape effects, a specific shape) by 1-based index.
/// </summary>
/// <remarks>
/// PowerPoint's animation object model does not have a separate "Emphasis" category — the same
/// <c>MsoAnimEffect</c> id can be used as an entrance effect (default), and the same id can be
/// flagged as an exit effect via <c>Effect.Exit</c>. This surface follows that model: callers pick
/// an effect by name and optionally mark it <c>isExit</c>; "emphasis" usage is simply adding an
/// effect to a shape that's already visible, with no distinct flag required by the COM API.
/// </remarks>
[ServiceCategory("animation", "Animation")]
[McpTool("animation", Title = "Animation Operations", Destructive = true, Category = "content",
    Description = "Add or delete shape entrance/emphasis/exit animation effects and read or set slide transitions in an open presentation session.")]
public interface IAnimationCommands
{
    /// <summary>
    /// Adds an animation effect to the shape at <paramref name="shapeIndex"/> on the given slide,
    /// identified by its <c>MsoAnimEffect</c> enum member name (e.g. <c>"msoAnimEffectFade"</c>,
    /// <c>"msoAnimEffectFly"</c>, <c>"msoAnimEffectZoom"</c>). See <c>animations.md</c> for the
    /// full supported name list. The effect is appended to the slide's
    /// <c>TimeLine.MainSequence</c>.
    /// </summary>
    /// <param name="isExit">
    /// When true, the effect is applied as the shape leaving the slide (exit) rather than the
    /// default entrance/emphasis behavior.
    /// </param>
    /// <param name="trigger">
    /// When the effect starts: <c>"on-click"</c> (default), <c>"with-previous"</c>, or
    /// <c>"after-previous"</c>.
    /// </param>
    AnimationOperationResult AddEffect(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string effectName,
        bool isExit = false,
        string trigger = "on-click");

    /// <summary>Gets the number of animation effects in the given slide's MainSequence timeline.</summary>
    AnimationOperationResult GetEffectCount(ComInterop.Session.IPresentationBatch batch, int slideIndex);

    /// <summary>Deletes the effect at the given 1-based index from the given slide's MainSequence timeline.</summary>
    AnimationOperationResult DeleteEffect(ComInterop.Session.IPresentationBatch batch, int slideIndex, int effectIndex);

    /// <summary>Gets the given slide's transition settings.</summary>
    AnimationOperationResult GetTransition(ComInterop.Session.IPresentationBatch batch, int slideIndex);

    /// <summary>
    /// Sets the given slide's transition, identified by its <c>PpEntryEffect</c> enum member name
    /// (e.g. <c>"ppEffectFade"</c>, <c>"ppEffectCut"</c>, <c>"ppEffectPushLeft"</c>). See
    /// <c>animations.md</c> for the full supported name list. Any optional parameter left null is
    /// unchanged.
    /// </summary>
    AnimationOperationResult SetTransition(
        ComInterop.Session.IPresentationBatch batch,
        int slideIndex,
        string transitionName,
        float? durationSeconds = null,
        bool? advanceOnClick = null,
        bool? advanceOnTime = null,
        float? advanceTimeSeconds = null);
}
