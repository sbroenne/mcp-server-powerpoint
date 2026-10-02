# Animations

Reference for `animation(action: "...", ...)` — adds/removes shape entrance/emphasis/exit
animation effects on a slide's timeline, and reads/sets a slide's transition (the effect that
plays when advancing *to* that slide during a slide show).

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `animation` | `add-effect` | `session_id`, `slide_index`, `shape_index`, `effect_name`, `is_exit?`, `trigger?` | Adds an effect to the shape's slide, appended to the slide's animation timeline. `is_exit` (default `false`) makes it play as the shape leaving rather than entering/emphasizing. `trigger` is `"on-click"` (default), `"with-previous"`, or `"after-previous"`. Returns `effect_index` (1-based, position in the timeline) and `effect_count`. |
| `animation` | `get-effect-count` | `session_id`, `slide_index` | Returns `effect_count` — the number of animation effects on the slide's timeline. |
| `animation` | `delete-effect` | `session_id`, `slide_index`, `effect_index` | Removes the effect at the given 1-based timeline position. |
| `animation` | `get-transition` | `session_id`, `slide_index` | Returns `transition_name`, `duration_seconds`, `advance_on_click`, `advance_on_time`, `advance_time_seconds` for the slide. |
| `animation` | `set-transition` | `session_id`, `slide_index`, `transition_name`, `duration_seconds?`, `advance_on_click?`, `advance_on_time?`, `advance_time_seconds?` | Sets the slide's transition effect and (optionally) its timing/advance behavior. Every parameter besides `transition_name` is optional — omit any you don't want to change. |

## Entrance/Emphasis/Exit Effect Names (`effect_name`)

PowerPoint's animation model doesn't have a separate "emphasis" flag — the same effect id is used
whether the shape is already visible (emphasis) or newly appearing (entrance); `is_exit` is the
only flag that changes behavior (shape leaving vs. entering/emphasizing). Supported `effect_name`
values (a curated subset of the full `MsoAnimEffect` enum):

| Name | Description |
|------|-------------|
| `msoAnimEffectAppear` | Appears instantly |
| `msoAnimEffectFly` | Flies in/out |
| `msoAnimEffectBlinds` | Blinds |
| `msoAnimEffectBox` | Box |
| `msoAnimEffectCheckerboard` | Checkerboard |
| `msoAnimEffectCircle` | Circle |
| `msoAnimEffectDiamond` | Diamond |
| `msoAnimEffectDissolve` | Dissolve |
| `msoAnimEffectFade` | Fade |
| `msoAnimEffectFlashOnce` | Flash once |
| `msoAnimEffectPeek` | Peek |
| `msoAnimEffectPlus` | Plus |
| `msoAnimEffectRandomBars` | Random bars |
| `msoAnimEffectSpiral` | Spiral |
| `msoAnimEffectSplit` | Split |
| `msoAnimEffectStretch` | Stretch |
| `msoAnimEffectStrips` | Strips |
| `msoAnimEffectSwivel` | Swivel |
| `msoAnimEffectWedge` | Wedge |
| `msoAnimEffectWheel` | Wheel |
| `msoAnimEffectWipe` | Wipe |
| `msoAnimEffectZoom` | Zoom |
| `msoAnimEffectBounce` | Bounce |
| `msoAnimEffectCredits` | Credits (rolling) |
| `msoAnimEffectGrowShrink` | Grow/Shrink (emphasis) |
| `msoAnimEffectSpin` | Spin (emphasis) |
| `msoAnimEffectTransparency` | Transparency change (emphasis) |
| `msoAnimEffectChangeFillColor` | Fill color change (emphasis) |
| `msoAnimEffectChangeFontColor` | Font color change (emphasis) |

## Slide Transition Names (`transition_name`)

A curated subset of the full `PpEntryEffect` enum:

| Name | Description |
|------|-------------|
| `ppEffectNone` | No transition |
| `ppEffectCut` | Cut |
| `ppEffectCutThroughBlack` | Cut through black |
| `ppEffectRandom` | Random transition |
| `ppEffectBlindsHorizontal` / `ppEffectBlindsVertical` | Blinds |
| `ppEffectCheckerboardAcross` / `ppEffectCheckerboardDown` | Checkerboard |
| `ppEffectCoverLeft` / `ppEffectCoverRight` / `ppEffectCoverUp` / `ppEffectCoverDown` | Cover |
| `ppEffectUncoverLeft` / `ppEffectUncoverRight` / `ppEffectUncoverUp` / `ppEffectUncoverDown` | Uncover |
| `ppEffectDissolve` | Dissolve |
| `ppEffectFade` | Fade |
| `ppEffectFadeSmoothly` | Fade smoothly |
| `ppEffectRandomBarsHorizontal` / `ppEffectRandomBarsVertical` | Random bars |
| `ppEffectPushLeft` / `ppEffectPushRight` / `ppEffectPushUp` / `ppEffectPushDown` | Push |
| `ppEffectWedge` | Wedge |
| `ppEffectCircleOut` | Circle |
| `ppEffectDiamondOut` | Diamond |
| `ppEffectPlusOut` | Plus |
| `ppEffectNewsflash` | Newsflash |

## Typical Use

```
1. shape.add-text-box(session_id, slide_index, ...) → shape_index
2. animation(action: "add-effect", session_id, slide_index, shape_index, effect_name: "msoAnimEffectFade", trigger: "on-click")
3. animation(action: "set-transition", session_id, slide_index, transition_name: "ppEffectFade", duration_seconds: 0.75)
```

Entrance effects are added per-shape via `add-effect`; transitions are set per-slide via
`set-transition` and control how the slide show moves *into* that slide.

## Extending the Curated Name Lists

Both `effect_name` and `transition_name` are validated against a curated subset of PowerPoint's
full (150+ member) `MsoAnimEffect`/`PpEntryEffect` enums — not every enum member is wired up.
Extending either list requires adding a verified numeric value to the corresponding
`Dictionary<string, int>` in `AnimationCommands.cs` (never guess a value; check
learn.microsoft.com's `office.vba` reference pages for the enum first).
