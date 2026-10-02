# Composition Recipes

Four editable slide patterns built with existing shape, textframe, chart, and notes
commands. Choose a pattern for the content, then export and inspect the result.
These are constrained examples, not automatic layout or overflow correction.

## Shared Geometry

Read `pagesetup(action: "get-settings", session_id: ...)` first. Let `W` and `H`
be the slide dimensions in points. The examples use `H=540` and `W=960` (16:9)
or `W=720` (4:3). Horizontal coordinates below are fractions of `W`; vertical
coordinates and font sizes are points at `H=540`. For a different height, multiply
vertical coordinates, heights, and font sizes by `H/540`, then verify the render.
Only the two example sizes have been tested.

- Start with `slide(action: "add-blank", ...)`. Use its returned slide index.
- Content starts at `0.055W`; its width is `0.89W`.
- Title box: left `0.055W`, top `40`, width `0.89W`, height `75`, 32pt bold.
- Examples use Aptos, dark text `RGB(28,35,40)`, and teal rules `RGB(0,115,110)`
  on white. For a branded deck, use its approved palette and verify
  foreground/background contrast.
- Add text with `shape(action: "add-text-box", ...)`, then use the returned
  shape index for `textframe` font and color actions. Rules are rectangles with
  a solid fill and no visible line. All elements remain editable.
- Add speaker notes, export the slide to PNG, and inspect it at its intended size.
  Check chart labels separately: text-box bounds do not validate chart internals.

## Comparison

Use for two alternatives evaluated on the same three criteria. Do not imply one
option is preferred through unequal widths or inconsistent ordering.

| Element | Left | Top | Width | Height | Type |
|---------|------|-----|-------|--------|------|
| Left/right heading | `0.055W` / `0.525W` | 150 | `0.42W` | 50 | 26pt bold |
| Left/right rule | same | 211 | `0.42W` | 4 | Filled rectangle |
| Left/right criteria | same | 245 | `0.42W` | 175 | 22pt, three short lines |
| Decision footer | `0.055W` | 455 | `0.89W` | 45 | 18pt |

Example: **Shared service** vs **Dedicated team**, with Start, Capacity, and Best
for in that order on both sides. Keep headings to a few words and each criterion
to a short phrase. More alternatives or long paragraphs belong on another slide
or in a table. At 4:3, rewrite long labels before reducing the body font.

## Chart and Insight

Use for one numeric series and one supported takeaway. Keep the chart native,
not a screenshot or a collection of hand-drawn bars.

| Element | Left | Top | Width | Height | Type |
|---------|------|-----|-------|--------|------|
| Native chart | `0.055W` | 145 | `0.57W` | 285 | `chart` add-chart |
| Takeaway | `0.65W` | 165 | `0.295W` | 75 | 28pt bold |
| Interpretation | `0.65W` | 255 | `0.295W` | 150 | 22pt |
| Source footer | `0.055W` | 455 | `0.89W` | 45 | 16pt |

Tested example: `chart_type: "bar"`, categories Q1-Q4, series Active teams,
values `[12,18,25,34]`, chart title Active teams, legend hidden. The takeaway is
**+22 teams**, the difference between Q4 and Q1, not a percentage. Interpretation:
"Growth continues. Next: verify retention." The source explicitly identifies
the data as synthetic.

Use short category labels and no more than four categories for this recipe.
More series or dense labels need a full-width chart. Preserve units, relevant
baseline, and source context; do not invent a causal explanation from a trend.

## Three-Step Timeline

Use for three ordered milestones, not a duration-scaled project schedule.

For step `i=0,1,2`, left is `0.055W + i*0.305W`; width is `0.28W`.

| Element | Top | Height | Type |
|---------|-----|--------|------|
| Period | 160 | 60 | 22pt |
| Rule | 235 | 5 | Filled rectangle |
| Milestone name | 265 | 60 | 26pt bold |
| Outcome | 335 | 100 | 20pt |

Example: Weeks 1-2 / Discover / Confirm the need; Weeks 3-4 / Pilot / Validate
with users; Week 5 / Launch / Release and monitor. Use one short outcome per
milestone. Add a 16pt footer at top `455`, height `45`, explaining that spacing
is not elapsed time. Dependencies, overlapping phases, or more milestones need
a different layout; equal spacing would misrepresent a true schedule.

## Metric Callout

Use for one value that needs a definition and a meaningful comparison.

| Element | Left | Top | Width | Height | Type |
|---------|------|-----|-------|--------|------|
| Value and unit | `0.055W` | 150 | `0.89W` | 120 | 78pt bold |
| Definition | same | 290 | same | 65 | 28pt |
| Rule | same | 380 | same | 4 | Filled rectangle |
| Baseline and cohort | same | 410 | same | 70 | 20pt, two lines |

Example: **18 days**, Median time to first result; Previously 24 days; Synthetic
cohort: 40 teams, last quarter. Do not omit the unit, denominator, time period,
or qualification needed to interpret the number. Multiple competing metrics or
a long qualification need a different slide, not several oversized numbers.

## Reproduce the Examples

`CompositionRecipeTests.Recipes_RenderWithContentInsideSlideBounds` in the Core
test project builds all four examples through public commands at both sizes.
It verifies text read-back, COM text bounds, chart category count, and PNG export.
All labels and data are authored synthetic examples; no external deck is needed.

On Windows with PowerPoint installed, run from the repository root:

```powershell
$env:PPTMCP_RECIPE_OUTPUT = Join-Path $env:TEMP 'powerpoint-composition-recipes'
try {
    dotnet test tests/PowerPointMcp.Core.Tests -c Release --filter 'FullyQualifiedName~CompositionRecipeTests' --blame-hang-timeout 5m
} finally {
    Remove-Item Env:PPTMCP_RECIPE_OUTPUT
}
```

The output contains `720x540` and `960x540` folders, each with `recipe-1.png`
through `recipe-4.png` in the order above. Without the environment variable,
the test deletes its temporary exports. Inspect all eight images after changing
fonts, content, or geometry. Passing bounds checks alone is not a visual review.