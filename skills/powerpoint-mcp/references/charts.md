# Charts

Reference for `chart(action: "add-chart", ...)` and `chart(action: "get-chart-data", ...)` —
native PowerPoint charts (not images of charts), backed by an embedded chart data sheet.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `chart` | `add-chart` | `session_id`, `slide_index`, `chart_type`, `left`, `top`, `width`, `height`, `categories` (string array), `series_name` (string), `values` (double array) | Creates a native chart shape with **one** data series. |
| `chart` | `get-chart-data` | `session_id`, `slide_index`, `shape_index` | Returns `categoryCount` and `seriesCount` of an existing chart — dimensions only, not the raw values. |

## Supported Chart Types

`chart_type` is a plain string: `"bar"`, `"line"`, or `"pie"`. There is no doughnut, scatter, area,
or 3D variant in this tool surface — pick the closest of the three:

| Need | Use |
|------|-----|
| Comparing categories side by side | `"bar"` |
| Trend over time | `"line"` |
| Part-of-whole (few categories) | `"pie"` |

## Single Series Only

`chart(action: "add-chart", ...)` takes exactly one `series_name` + one `values` array — there is
no multi-series parameter set in this tool surface. `categories` and `values` must be the same
length (one value per category label):

```
chart(action: "add-chart", session_id: ..., slide_index: ..., chart_type: "bar",
  left: 60, top: 120, width: 500, height: 300,
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series_name: "Revenue",
  values: [120.0, 150.0, 170.0, 210.0])
```

If a task genuinely needs multiple series (e.g., "Revenue" and "Cost" side by side), the current
surface cannot combine them into one chart object — either pick the single most important series,
or place two separate `chart(action: "add-chart", ...)` calls side-by-side on the slide (each with
its own categories) and label them clearly with adjacent text boxes.

## Pie Charts

For `"pie"`, `categories` become the slice labels and `values` the slice sizes — keep to 6 or
fewer categories so labels stay legible once rendered.

## Sizing and Placement

- Keep charts within the slide's safe area (see `deck-builder.md` positioning reference).
- Leave room beside or below the chart for a text-box callout describing the key takeaway — this
  tool surface has no chart title/axis-label parameters, so a nearby `shape(action:
  "add-text-box", ...)` is how you label the insight.
- Minimum practical size: `width ≥ 300, height ≥ 200` — smaller charts render illegibly once
  exported.

## Reading Back Chart Data

`chart(action: "get-chart-data", ...)` only reports `categoryCount`/`seriesCount` — it does not
return the actual category labels or values. Use it to confirm a chart was created with the
expected shape (e.g., 4 categories, 1 series) after `add-chart`, not to recover the original data
for editing — there is no set/edit-in-place action; to change chart contents, delete the shape
(`shape(action: "delete", ...)`) and call `add-chart` again with corrected data.

## Verify Visually

Charts are the highest-value target for `export(action: "export-slide-to-image", ...)` —
data-entry mistakes (wrong values, mismatched category count) are invisible from a
`get-chart-data` call alone but obvious in the rendered image. Always export and inspect after
`add-chart` (see `export-and-verify.md`).
