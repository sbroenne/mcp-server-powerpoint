# Charts

Reference for the `chart` tool — native PowerPoint charts (not images of charts), backed by an
embedded chart data sheet.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `chart` | `add-chart` | `session_id`, `slide_index`, `chart_type`, `left`, `top`, `width`, `height`, `categories` (string array), `series_name` (string), `values` (double array) | Creates a native chart shape with **one** data series. |
| `chart` | `get-chart-data` | `session_id`, `slide_index`, `shape_index` | Returns `categoryCount` and `seriesCount` of an existing chart — dimensions only, not the raw values. |
| `chart` | `add-series` | `session_id`, `slide_index`, `shape_index`, `series_name`, `values` (double array) | Adds one more data series to an existing chart. `values` length must match the chart's existing category count. Call repeatedly to build an N-series chart. |
| `chart` | `replace-chart-data` | `session_id`, `slide_index`, `shape_index`, `categories` (string array), `series_names` (string array), `series_values` (double array, **series-major flat**) | Replaces ALL of an existing chart's categories/series/values in one call — including changing the category count. See "Replacing Chart Data" below for the flat layout. |
| `chart` | `set-chart-title` | `session_id`, `slide_index`, `shape_index`, `title` | Sets/shows the chart's title text. |
| `chart` | `get-chart-title` | `session_id`, `slide_index`, `shape_index` | Returns `hasTitle` and, if present, `title`. |
| `chart` | `set-axis-title` | `session_id`, `slide_index`, `shape_index`, `axis_type` (`"category"` or `"value"`), `title` | Sets the title of the category (X) or value (Y) axis. |
| `chart` | `get-axis-title` | `session_id`, `slide_index`, `shape_index`, `axis_type` (`"category"` or `"value"`) | Returns the axis title text. |
| `chart` | `set-legend-visibility` | `session_id`, `slide_index`, `shape_index`, `visible` (bool) | Shows/hides the chart's legend. |
| `chart` | `get-legend-visibility` | `session_id`, `slide_index`, `shape_index` | Returns `legendVisible`. |

## Supported Chart Types

`chart_type` is a plain string: `"bar"`, `"line"`, or `"pie"`. There is no doughnut, scatter, area,
or 3D variant in this tool surface — pick the closest of the three:

| Need | Use |
|------|-----|
| Comparing categories side by side | `"bar"` |
| Trend over time | `"line"` |
| Part-of-whole (few categories) | `"pie"` |

## Multi-Series Charts

`chart(action: "add-chart", ...)` always creates the chart with exactly one series. To add more
series (e.g., "Revenue" and "Cost" side by side), call `chart(action: "add-series", ...)` once per
additional series against the shape returned by `add-chart`:

```
chart(action: "add-chart", session_id: ..., slide_index: ..., chart_type: "bar",
  left: 60, top: 120, width: 500, height: 300,
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series_name: "Revenue",
  values: [120.0, 150.0, 170.0, 210.0])
# → shapeIndex from the result above

chart(action: "add-series", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  series_name: "Cost", values: [80.0, 95.0, 110.0, 130.0])
```

Each `add-series` call's `values` array length must match the chart's existing category count
(from the original `add-chart` call) — a mismatch returns `Success=false` without throwing.

## Replacing Chart Data

Unlike `add-series` (which appends one more series to the existing category count),
`chart(action: "replace-chart-data", ...)` wholesale-replaces an existing chart's categories, series
names, and values in a single call — including changing the number of categories. This avoids the
delete-shape-and-recreate workaround mentioned above.

`series_values` is a **flat, series-major** array: all values for `series_names[0]` first, then all
values for `series_names[1]`, etc. Its length must equal `categories.length * series_names.length`.

```
chart(action: "replace-chart-data", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  categories: ["Jan", "Feb", "Mar", "Apr"],
  series_names: ["Revenue", "Cost"],
  # Revenue: 100, 200, 300, 400 — then Cost: 50, 60, 70, 80
  series_values: [100.0, 200.0, 300.0, 400.0, 50.0, 60.0, 70.0, 80.0])
```

A mismatched `series_values` length or an invalid/non-chart `shape_index` returns `Success=false`
without throwing.

## Pie Charts

For `"pie"`, `categories` become the slice labels and `values` the slice sizes — keep to 6 or
fewer categories so labels stay legible once rendered. `add-series` is rarely useful on a pie
chart (pie charts render only their first series).

## Titles and Legend

Use `set-chart-title`/`set-axis-title` to label the chart and its axes directly, instead of (or in
addition to) a nearby text-box callout:

```
chart(action: "set-chart-title", session_id: ..., slide_index: ..., shape_index: ..., title: "Quarterly Revenue")
chart(action: "set-axis-title", session_id: ..., slide_index: ..., shape_index: ..., axis_type: "category", title: "Quarter")
chart(action: "set-axis-title", session_id: ..., slide_index: ..., shape_index: ..., axis_type: "value", title: "USD (thousands)")
chart(action: "set-legend-visibility", session_id: ..., slide_index: ..., shape_index: ..., visible: true)
```

`set-legend-visibility` with `visible: true` is recommended whenever a chart has more than one
series — without a visible legend, a multi-series chart's colors are unlabeled.

## Sizing and Placement

- Keep charts within the slide's safe area (see `deck-builder.md` positioning reference).
- Leave room beside or below the chart for a text-box callout describing the key takeaway, unless
  you've already used `set-chart-title`/`set-axis-title` to label it directly.
- Minimum practical size: `width ≥ 300, height ≥ 200` — smaller charts render illegibly once
  exported.

## Reading Back Chart Data

`chart(action: "get-chart-data", ...)` only reports `categoryCount`/`seriesCount` — it does not
return the actual category labels or values. Use it to confirm a chart was created with the
expected shape (e.g., 4 categories, 1 series) after `add-chart`, not to recover the original data
for editing. To change the category labels or values, use `replace-chart-data` (see "Replacing
Chart Data" above) rather than deleting and recreating the shape.

## Verify Visually

Charts are the highest-value target for `export(action: "export-slide-to-image", ...)` —
data-entry mistakes (wrong values, mismatched category count) are invisible from a
`get-chart-data` call alone but obvious in the rendered image. Always export and inspect after
`add-chart`/`add-series` (see `export-and-verify.md`).

