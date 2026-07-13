using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Animation;

/// <inheritdoc cref="IAnimationCommands"/>
public sealed class AnimationCommands : IAnimationCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // MsoAnimTriggerType (PowerPoint namespace, embeddable without office.dll).
    private static readonly Dictionary<string, PowerPoint.MsoAnimTriggerType> TriggerTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["on-click"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick,
        ["with-previous"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
        ["after-previous"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious,
    };

    // MsoAnimEffect (PowerPoint namespace, embeddable without office.dll). Curated subset of the
    // full (150+) enum covering the
    // classic entrance/emphasis/exit effects authors most commonly need — verified against
    // learn.microsoft.com/office/vba/api/powerpoint.msoanimeffect. Extend this table (never guess
    // a value) if more effects are needed later.
    private static readonly Dictionary<string, PowerPoint.MsoAnimEffect> AnimEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoAnimEffectAppear"] = PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
        ["msoAnimEffectFly"] = PowerPoint.MsoAnimEffect.msoAnimEffectFly,
        ["msoAnimEffectBlinds"] = PowerPoint.MsoAnimEffect.msoAnimEffectBlinds,
        ["msoAnimEffectBox"] = PowerPoint.MsoAnimEffect.msoAnimEffectBox,
        ["msoAnimEffectCheckerboard"] = PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard,
        ["msoAnimEffectCircle"] = PowerPoint.MsoAnimEffect.msoAnimEffectCircle,
        ["msoAnimEffectDiamond"] = PowerPoint.MsoAnimEffect.msoAnimEffectDiamond,
        ["msoAnimEffectDissolve"] = PowerPoint.MsoAnimEffect.msoAnimEffectDissolve,
        ["msoAnimEffectFade"] = PowerPoint.MsoAnimEffect.msoAnimEffectFade,
        ["msoAnimEffectFlashOnce"] = PowerPoint.MsoAnimEffect.msoAnimEffectFlashOnce,
        ["msoAnimEffectPeek"] = PowerPoint.MsoAnimEffect.msoAnimEffectPeek,
        ["msoAnimEffectPlus"] = PowerPoint.MsoAnimEffect.msoAnimEffectPlus,
        ["msoAnimEffectRandomBars"] = PowerPoint.MsoAnimEffect.msoAnimEffectRandomBars,
        ["msoAnimEffectSpiral"] = PowerPoint.MsoAnimEffect.msoAnimEffectSpiral,
        ["msoAnimEffectSplit"] = PowerPoint.MsoAnimEffect.msoAnimEffectSplit,
        ["msoAnimEffectStretch"] = PowerPoint.MsoAnimEffect.msoAnimEffectStretch,
        ["msoAnimEffectStrips"] = PowerPoint.MsoAnimEffect.msoAnimEffectStrips,
        ["msoAnimEffectSwivel"] = PowerPoint.MsoAnimEffect.msoAnimEffectSwivel,
        ["msoAnimEffectWedge"] = PowerPoint.MsoAnimEffect.msoAnimEffectWedge,
        ["msoAnimEffectWheel"] = PowerPoint.MsoAnimEffect.msoAnimEffectWheel,
        ["msoAnimEffectWipe"] = PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
        ["msoAnimEffectZoom"] = PowerPoint.MsoAnimEffect.msoAnimEffectZoom,
        ["msoAnimEffectBounce"] = PowerPoint.MsoAnimEffect.msoAnimEffectBounce,
        ["msoAnimEffectCredits"] = PowerPoint.MsoAnimEffect.msoAnimEffectCredits,
        ["msoAnimEffectGrowShrink"] = PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
        ["msoAnimEffectSpin"] = PowerPoint.MsoAnimEffect.msoAnimEffectSpin,
        ["msoAnimEffectTransparency"] = PowerPoint.MsoAnimEffect.msoAnimEffectTransparency,
        ["msoAnimEffectChangeFillColor"] = PowerPoint.MsoAnimEffect.msoAnimEffectChangeFillColor,
        ["msoAnimEffectChangeFontColor"] = PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor,
    };

    // PpEntryEffect member name -> value, a curated subset of the full (150+) enum covering the
    // most common slide transitions — verified against
    // learn.microsoft.com/office/vba/api/powerpoint.ppentryeffect. Extend this table (never guess
    // a value) if more transitions are needed later.
    private static readonly Dictionary<string, PowerPoint.PpEntryEffect> EntryEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ppEffectNone"] = PowerPoint.PpEntryEffect.ppEffectNone,
        ["ppEffectCut"] = PowerPoint.PpEntryEffect.ppEffectCut,
        ["ppEffectCutThroughBlack"] = PowerPoint.PpEntryEffect.ppEffectCutThroughBlack,
        ["ppEffectRandom"] = PowerPoint.PpEntryEffect.ppEffectRandom,
        ["ppEffectBlindsHorizontal"] = PowerPoint.PpEntryEffect.ppEffectBlindsHorizontal,
        ["ppEffectBlindsVertical"] = PowerPoint.PpEntryEffect.ppEffectBlindsVertical,
        ["ppEffectCheckerboardAcross"] = PowerPoint.PpEntryEffect.ppEffectCheckerboardAcross,
        ["ppEffectCheckerboardDown"] = PowerPoint.PpEntryEffect.ppEffectCheckerboardDown,
        ["ppEffectCoverLeft"] = PowerPoint.PpEntryEffect.ppEffectCoverLeft,
        ["ppEffectCoverUp"] = PowerPoint.PpEntryEffect.ppEffectCoverUp,
        ["ppEffectCoverRight"] = PowerPoint.PpEntryEffect.ppEffectCoverRight,
        ["ppEffectCoverDown"] = PowerPoint.PpEntryEffect.ppEffectCoverDown,
        ["ppEffectDissolve"] = PowerPoint.PpEntryEffect.ppEffectDissolve,
        ["ppEffectFade"] = PowerPoint.PpEntryEffect.ppEffectFade,
        ["ppEffectRandomBarsHorizontal"] = PowerPoint.PpEntryEffect.ppEffectRandomBarsHorizontal,
        ["ppEffectRandomBarsVertical"] = PowerPoint.PpEntryEffect.ppEffectRandomBarsVertical,
        ["ppEffectUncoverLeft"] = PowerPoint.PpEntryEffect.ppEffectUncoverLeft,
        ["ppEffectUncoverUp"] = PowerPoint.PpEntryEffect.ppEffectUncoverUp,
        ["ppEffectUncoverRight"] = PowerPoint.PpEntryEffect.ppEffectUncoverRight,
        ["ppEffectUncoverDown"] = PowerPoint.PpEntryEffect.ppEffectUncoverDown,
        ["ppEffectWedge"] = PowerPoint.PpEntryEffect.ppEffectWedge,
        ["ppEffectCircleOut"] = PowerPoint.PpEntryEffect.ppEffectCircleOut,
        ["ppEffectDiamondOut"] = PowerPoint.PpEntryEffect.ppEffectDiamondOut,
        ["ppEffectPlusOut"] = PowerPoint.PpEntryEffect.ppEffectPlusOut,
        ["ppEffectPushLeft"] = PowerPoint.PpEntryEffect.ppEffectPushLeft,
        ["ppEffectPushRight"] = PowerPoint.PpEntryEffect.ppEffectPushRight,
        ["ppEffectPushUp"] = PowerPoint.PpEntryEffect.ppEffectPushUp,
        ["ppEffectPushDown"] = PowerPoint.PpEntryEffect.ppEffectPushDown,
        ["ppEffectNewsflash"] = PowerPoint.PpEntryEffect.ppEffectNewsflash,
        ["ppEffectFadeSmoothly"] = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly,
    };

    private static readonly Dictionary<PowerPoint.PpEntryEffect, string> EntryEffectsByValue =
        EntryEffects.GroupBy(kvp => kvp.Value).ToDictionary(g => g.Key, g => g.First().Key);

    /// <inheritdoc/>
    public AnimationOperationResult AddEffect(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string effectName,
        bool isExit = false,
        string trigger = "on-click")
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(effectName);
        ArgumentNullException.ThrowIfNull(trigger);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            if (!AnimEffects.TryGetValue(effectName, out var effectValue))
            {
                return new AnimationOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{effectName}' is not a recognized MsoAnimEffect name (e.g. 'msoAnimEffectFade', 'msoAnimEffectFly', 'msoAnimEffectZoom')."
                };
            }

            if (!TriggerTypes.TryGetValue(trigger, out var triggerValue))
            {
                return new AnimationOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{trigger}' is not a recognized trigger (must be 'on-click', 'with-previous', or 'after-previous')."
                };
            }

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            // Pre-capture append position: Count+1 before the call equals Count after (the new
            // effect's 1-based index). Passing 0 explicitly sends an integer-out-of-range COM
            // error — the VBA "default=0 means append" only works via COM's missing-arg protocol,
            // not from an explicit .NET value. Any index > Count is documented to append.
            int newIndex = sequence.Count + 1;
            PowerPoint.Effect effect = sequence.AddEffect(
                shape,
                effectValue,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                triggerValue,
                newIndex);
            dynamic? dynEffect = null;
            try
            {
                dynEffect = effect;
                dynEffect.Exit = isExit ? MsoTrue : MsoFalse;
                effect.Timing.TriggerType = triggerValue;

                return new AnimationOperationResult
                {
                    Success = true,
                    EffectIndex = newIndex,
                    EffectCount = sequence.Count,
                    EffectName = effectName,
                    IsExit = isExit,
                    Trigger = trigger
                };
            }
            finally
            {
                if (dynEffect != null)
                {
                    ComUtilities.Release(ref dynEffect!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public AnimationOperationResult GetEffectCount(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;

            return new AnimationOperationResult { Success = true, EffectCount = sequence.Count };
        });
    }

    /// <inheritdoc/>
    public AnimationOperationResult DeleteEffect(IPresentationBatch batch, int slideIndex, int effectIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            int effectCount = sequence.Count;

            if (effectIndex < 1 || effectIndex > effectCount)
            {
                return new AnimationOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Effect index {effectIndex} is out of range. The slide has {effectCount} effect(s) (valid range: 1-{effectCount})."
                };
            }

            PowerPoint.Effect effect = sequence[effectIndex];
            effect.Delete();

            return new AnimationOperationResult
            {
                Success = true,
                EffectIndex = effectIndex,
                EffectCount = sequence.Count
            };
        });
    }

    /// <inheritdoc/>
    public AnimationOperationResult GetTransition(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            return ReadTransition(slide);
        });
    }

    /// <inheritdoc/>
    public AnimationOperationResult SetTransition(
        IPresentationBatch batch,
        int slideIndex,
        string transitionName,
        float? durationSeconds = null,
        bool? advanceOnClick = null,
        bool? advanceOnTime = null,
        float? advanceTimeSeconds = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(transitionName);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            if (!EntryEffects.TryGetValue(transitionName, out var transitionValue))
            {
                return new AnimationOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{transitionName}' is not a recognized PpEntryEffect name (e.g. 'ppEffectFade', 'ppEffectCut', 'ppEffectPushLeft')."
                };
            }

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            PowerPoint.SlideShowTransition transition = slide.SlideShowTransition;
            dynamic? dynTransition = null;
            try
            {
                transition.EntryEffect = transitionValue;

                if (durationSeconds is not null)
                {
                    transition.Duration = durationSeconds.Value;
                }

                if (advanceOnClick is not null)
                {
                    dynTransition = transition;
                    dynTransition.AdvanceOnClick = advanceOnClick.Value ? MsoTrue : MsoFalse;
                }

                if (advanceOnTime is not null)
                {
                    dynTransition ??= transition;
                    dynTransition.AdvanceOnTime = advanceOnTime.Value ? MsoTrue : MsoFalse;
                }

                if (advanceTimeSeconds is not null)
                {
                    transition.AdvanceTime = advanceTimeSeconds.Value;
                }

                return ReadTransition(slide);
            }
            finally
            {
                if (dynTransition != null)
                {
                    ComUtilities.Release(ref dynTransition!);
                }
            }
        });
    }

    private static AnimationOperationResult ReadTransition(PowerPoint.Slide slide)
    {
        PowerPoint.SlideShowTransition transition = slide.SlideShowTransition;
        PowerPoint.PpEntryEffect entryEffectValue = transition.EntryEffect;
        string transitionName = EntryEffectsByValue.TryGetValue(entryEffectValue, out var name)
            ? name
            : $"unknown ({(int)entryEffectValue})";
        dynamic? dynTransition = null;
        try
        {
            dynTransition = transition;

            return new AnimationOperationResult
            {
                Success = true,
                TransitionName = transitionName,
                DurationSeconds = transition.Duration,
                AdvanceOnClick = (int)dynTransition.AdvanceOnClick == MsoTrue,
                AdvanceOnTime = (int)dynTransition.AdvanceOnTime == MsoTrue,
                AdvanceTimeSeconds = transition.AdvanceTime
            };
        }
        finally
        {
            if (dynTransition != null)
            {
                ComUtilities.Release(ref dynTransition!);
            }
        }
    }

    private static AnimationOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new AnimationOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }

    private static AnimationOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new AnimationOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }
        return null;
    }
}
