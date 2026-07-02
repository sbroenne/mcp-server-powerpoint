# Charts

Reference for `add_chart` and `get_chart_data` — native PowerPoint charts (not images of charts),
backed by an embedded chart data sheet.

## Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `add_chart` | `sessionId`, `slideIndex`, `chartType`, `left`, `top`, `width`, `height`, `categories` (string array), `seriesName` (string), `values` (double array) | Creates a native chart shape with **one** data series. |
| `get_chart_data` | `sessionId`, `slideIndex`, `shapeIndex` | Returns `categoryCount` and `seriesCount` of an existing chart — dimensions only, not the raw values. |

## Supported Chart Types

`chartType` is a plain string: `"bar"`, `"line"`, or `"pie"`. There is no doughnut, scatter, area,
or 3D variant in this tool surface — pick the closest of the three:

| Need | Use |
|------|-----|
| Comparing categories side by side | `"bar"` |
| Trend over time | `"line"` |
| Part-of-whole (few categories) | `"pie"` |

## Single Series Only

`add_chart` takes exactly one `seriesName` + one `values` array — there is no multi-series
parameter set in this tool surface. `categories` and `values` must be the same length (one value
per category label):

```
add_chart(sessionId, slideIndex, chartType="bar",
  left=60, top=120, width=500, height=300,
  categories=["Q1", "Q2", "Q3", "Q4"],
  seriesName="Revenue",
  values=[120.0, 150.0, 170.0, 210.0])
```

If a task genuinely needs multiple series (e.g., "Revenue" and "Cost" side by side), the current
surface cannot combine them into one chart object — either pick the single most important series,
or place two separate `add_chart` calls side-by-side on the slide (each with its own categories)
and label them clearly with adjacent text boxes.

## Pie Charts

For `"pie"`, `categories` become the slice labels and `values` the slice sizes — keep to 6 or
fewer categories so labels stay legible once rendered.

## Sizing and Placement

- Keep charts within the slide's safe area (see `deck-builder.md` positioning reference).
- Leave room beside or below the chart for a text-box callout describing the key takeaway — this
  tool surface has no chart title/axis-label parameters, so a nearby `add_text_box` is how you
  label the insight.
- Minimum practical size: `width ≥ 300, height ≥ 200` — smaller charts render illegibly once
  exported.

## Reading Back Chart Data

`get_chart_data` only reports `categoryCount`/`seriesCount` — it does not return the actual
category labels or values. Use it to confirm a chart was created with the expected shape (e.g.,
4 categories, 1 series) after `add_chart`, not to recover the original data for editing — there is
no `set_chart_data`/edit-in-place tool; to change chart contents, delete the shape
(`delete_shape`) and call `add_chart` again with corrected data.

## Verify Visually

Charts are the highest-value target for `export_slide_to_image` — data-entry mistakes (wrong
values, mismatched category count) are invisible from a `get_chart_data` call alone but obvious
in the rendered image. Always export and inspect after `add_chart` (see `export-and-verify.md`).
