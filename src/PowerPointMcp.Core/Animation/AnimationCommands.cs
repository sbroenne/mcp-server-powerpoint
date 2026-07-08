using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Animation;

/// <inheritdoc cref="IAnimationCommands"/>
public sealed class AnimationCommands : IAnimationCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // MsoAnimTriggerType member -> value (learn.microsoft.com/office/vba/api/powerpoint.msoanimtriggertype).
    private static readonly Dictionary<string, int> TriggerTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["on-click"] = 1,   // msoAnimTriggerOnPageClick
        ["with-previous"] = 2,   // msoAnimTriggerWithPrevious
        ["after-previous"] = 3,   // msoAnimTriggerAfterPrevious
    };

    // MsoAnimEffect member name -> value, a curated subset of the full (150+) enum covering the
    // classic entrance/emphasis/exit effects authors most commonly need — verified against
    // learn.microsoft.com/office/vba/api/powerpoint.msoanimeffect. Extend this table (never guess
    // a value) if more effects are needed later.
    private static readonly Dictionary<string, int> AnimEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoAnimEffectAppear"] = 1,
        ["msoAnimEffectFly"] = 2,
        ["msoAnimEffectBlinds"] = 3,
        ["msoAnimEffectBox"] = 4,
        ["msoAnimEffectCheckerboard"] = 5,
        ["msoAnimEffectCircle"] = 6,
        ["msoAnimEffectDiamond"] = 8,
        ["msoAnimEffectDissolve"] = 9,
        ["msoAnimEffectFade"] = 10,
        ["msoAnimEffectFlashOnce"] = 11,
        ["msoAnimEffectPeek"] = 12,
        ["msoAnimEffectPlus"] = 13,
        ["msoAnimEffectRandomBars"] = 14,
        ["msoAnimEffectSpiral"] = 15,
        ["msoAnimEffectSplit"] = 16,
        ["msoAnimEffectStretch"] = 17,
        ["msoAnimEffectStrips"] = 18,
        ["msoAnimEffectSwivel"] = 19,
        ["msoAnimEffectWedge"] = 20,
        ["msoAnimEffectWheel"] = 21,
        ["msoAnimEffectWipe"] = 22,
        ["msoAnimEffectZoom"] = 23,
        ["msoAnimEffectBounce"] = 26,
        ["msoAnimEffectCredits"] = 28,
        ["msoAnimEffectGrowShrink"] = 59,
        ["msoAnimEffectSpin"] = 61,
        ["msoAnimEffectTransparency"] = 62,
        ["msoAnimEffectChangeFillColor"] = 54,
        ["msoAnimEffectChangeFontColor"] = 56,
    };

    // PpEntryEffect member name -> value, a curated subset of the full (150+) enum covering the
    // most common slide transitions — verified against
    // learn.microsoft.com/office/vba/api/powerpoint.ppentryeffect. Extend this table (never guess
    // a value) if more transitions are needed later.
    private static readonly Dictionary<string, int> EntryEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ppEffectNone"] = 0,
        ["ppEffectCut"] = 257,
        ["ppEffectCutThroughBlack"] = 258,
        ["ppEffectRandom"] = 513,
        ["ppEffectBlindsHorizontal"] = 769,
        ["ppEffectBlindsVertical"] = 770,
        ["ppEffectCheckerboardAcross"] = 1025,
        ["ppEffectCheckerboardDown"] = 1026,
        ["ppEffectCoverLeft"] = 1281,
        ["ppEffectCoverUp"] = 1282,
        ["ppEffectCoverRight"] = 1283,
        ["ppEffectCoverDown"] = 1284,
        ["ppEffectDissolve"] = 1537,
        ["ppEffectFade"] = 1793,
        ["ppEffectRandomBarsHorizontal"] = 2305,
        ["ppEffectRandomBarsVertical"] = 2306,
        ["ppEffectUncoverLeft"] = 2049,
        ["ppEffectUncoverUp"] = 2050,
        ["ppEffectUncoverRight"] = 2051,
        ["ppEffectUncoverDown"] = 2052,
        ["ppEffectWedge"] = 3856,
        ["ppEffectCircleOut"] = 3845,
        ["ppEffectDiamondOut"] = 3846,
        ["ppEffectPlusOut"] = 3851,
        ["ppEffectPushLeft"] = 3853,
        ["ppEffectPushRight"] = 3854,
        ["ppEffectPushUp"] = 3852,
        ["ppEffectPushDown"] = 3855,
        ["ppEffectNewsflash"] = 3850,
        ["ppEffectFadeSmoothly"] = 3849,
    };

    private static readonly Dictionary<int, string> EntryEffectsByValue =
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
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

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic sequence = slide.TimeLine.MainSequence;
            dynamic effect = sequence.AddEffect(shape, effectValue);
            effect.Exit = isExit ? MsoTrue : MsoFalse;
            effect.Timing.TriggerType = triggerValue;

            // Same NoPIA late-binding quirk documented in ShapeCommands — the newly-added effect
            // is always appended, so its 1-based index is simply the new sequence Count.
            int newIndex = (int)sequence.Count;

            return new AnimationOperationResult
            {
                Success = true,
                EffectIndex = newIndex,
                EffectCount = (int)sequence.Count,
                EffectName = effectName,
                IsExit = isExit,
                Trigger = trigger
            };
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            dynamic sequence = slide.TimeLine.MainSequence;

            return new AnimationOperationResult { Success = true, EffectCount = (int)sequence.Count };
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            dynamic sequence = slide.TimeLine.MainSequence;
            int effectCount = (int)sequence.Count;

            if (effectIndex < 1 || effectIndex > effectCount)
            {
                return new AnimationOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Effect index {effectIndex} is out of range. The slide has {effectCount} effect(s) (valid range: 1-{effectCount})."
                };
            }

            dynamic effect = sequence[effectIndex];
            effect.Delete();

            return new AnimationOperationResult
            {
                Success = true,
                EffectIndex = effectIndex,
                EffectCount = (int)sequence.Count
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            dynamic transition = slide.SlideShowTransition;
            transition.EntryEffect = transitionValue;

            if (durationSeconds is not null)
            {
                transition.Duration = durationSeconds.Value;
            }

            if (advanceOnClick is not null)
            {
                transition.AdvanceOnClick = advanceOnClick.Value ? MsoTrue : MsoFalse;
            }

            if (advanceOnTime is not null)
            {
                transition.AdvanceOnTime = advanceOnTime.Value ? MsoTrue : MsoFalse;
            }

            if (advanceTimeSeconds is not null)
            {
                transition.AdvanceTime = advanceTimeSeconds.Value;
            }

            return ReadTransition(slide);
        });
    }

    private static AnimationOperationResult ReadTransition(dynamic slide)
    {
        dynamic transition = slide.SlideShowTransition;
        int entryEffectValue = (int)transition.EntryEffect;
        string transitionName = EntryEffectsByValue.TryGetValue(entryEffectValue, out var name)
            ? name
            : $"unknown ({entryEffectValue})";

        return new AnimationOperationResult
        {
            Success = true,
            TransitionName = transitionName,
            DurationSeconds = (float)transition.Duration,
            AdvanceOnClick = (int)transition.AdvanceOnClick == MsoTrue,
            AdvanceOnTime = (int)transition.AdvanceOnTime == MsoTrue,
            AdvanceTimeSeconds = (float)transition.AdvanceTime
        };
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
