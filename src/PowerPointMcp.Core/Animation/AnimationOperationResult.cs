namespace Sbroenne.PowerPointMcp.Core.Animation;

/// <summary>
/// Result of an animation operation (shape entrance/emphasis/exit effects, slide transitions).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as the other domain results (Rule 1):
/// Success == true implies ErrorMessage is null/empty.
/// </remarks>
public sealed class AnimationOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>1-based index of the effect within the slide's MainSequence timeline, if applicable.</summary>
    public int? EffectIndex { get; init; }

    /// <summary>Total effect count in the slide's MainSequence timeline after the operation, if applicable.</summary>
    public int? EffectCount { get; init; }

    /// <summary>The MsoAnimEffect name used for the effect (e.g. "msoAnimEffectFade"), if applicable.</summary>
    public string? EffectName { get; init; }

    /// <summary>Whether the effect is an exit effect (true) or an entrance/emphasis effect (false), if applicable.</summary>
    public bool? IsExit { get; init; }

    /// <summary>The trigger for the effect: "on-click", "with-previous", or "after-previous", if applicable.</summary>
    public string? Trigger { get; init; }

    /// <summary>The PpEntryEffect name used for the slide transition (e.g. "ppEffectFade"), if applicable.</summary>
    public string? TransitionName { get; init; }

    /// <summary>Slide transition duration in seconds, if applicable.</summary>
    public float? DurationSeconds { get; init; }

    /// <summary>Whether the slide advances on mouse click, if applicable.</summary>
    public bool? AdvanceOnClick { get; init; }

    /// <summary>Whether the slide advances automatically after a fixed time, if applicable.</summary>
    public bool? AdvanceOnTime { get; init; }

    /// <summary>Seconds after which the slide automatically advances, when <see cref="AdvanceOnTime"/> is true.</summary>
    public float? AdvanceTimeSeconds { get; init; }
}
