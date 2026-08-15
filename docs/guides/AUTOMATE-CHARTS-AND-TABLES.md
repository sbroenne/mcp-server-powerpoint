# Automate PowerPoint Charts and Tables

How to put real data on a slide as a **native** PowerPoint chart or table — not a screenshot of
one — from an AI assistant or a script.

## Native vs. pasted

A native chart keeps its own embedded data sheet. Someone opening the deck can click it, edit the
numbers, restyle it, or change the chart type. A pasted image cannot be edited, does not respond
to the theme, and looks wrong the moment the deck is resized.

Everything below produces native objects, because PowerPoint itself creates them.

## Charts

### Create a chart

```
chart(action: "add-chart", session_id: ..., slide_index: 4, chart_type: "bar",
      left: 60, top: 120, width: 500, height: 300,
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series_name: "Revenue",
      values: [120.0, 150.0, 170.0, 210.0])
```

`chart_type` is one of `"bar"`, `"line"` or `"pie"`:

| Need | Use |
|------|-----|
| Comparing categories side by side | `"bar"` |
| Trend over time | `"line"` |
| Part of a whole, with few slices | `"pie"` |

Positions and sizes are in points (1 inch = 72 points).

### Add more series

`add-chart` always creates the chart with exactly one series. Add the rest against the
`shapeIndex` it returned:

```
chart(action: "add-series", session_id: ..., slide_index: 4, shape_index: 2,
      series_name: "Cost", values: [80.0, 95.0, 110.0, 130.0])
```

Each series must supply exactly as many values as the chart has categories. A mismatch is
rejected cleanly with `success: false` rather than throwing.

### Replace the data later

To refresh a chart in an existing deck — new quarter, new numbers, possibly a different number of
categories — use one call rather than deleting and recreating the shape:

```
chart(action: "replace-chart-data", session_id: ..., slide_index: 4, shape_index: 2,
      categories: ["Q1", "Q2", "Q3", "Q4", "Q1"],
      series_names: ["Revenue", "Cost"],
      series_values: [120.0, 150.0, 170.0, 210.0, 230.0,
                       80.0,  95.0, 110.0, 130.0, 140.0])
```

`series_values` is a **flat, series-major** array: every value for the first series, then every
value for the second. Its length must equal `len(series_names) * len(categories)`.

Recreating the shape instead would lose its position, size and any formatting applied to it.

### Titles, axes and legend

```
chart(action: "set-chart-title", ..., title: "Quarterly revenue")
chart(action: "set-axis-title", ..., axis_type: "value", title: "USD millions")
chart(action: "set-legend-visibility", ..., visible: true)
```

Turn the legend off for single-series charts — it adds nothing and eats horizontal space.

## Tables

### Create and fill

```
table(action: "add-table", session_id: ..., slide_index: 5,
      rows: 4, columns: 3, left: 60, top: 130, width: 800, height: 240)
# → shapeIndex

table(action: "set-cell-text", ..., shape_index: 2, row: 1, column: 1, text: "Region")
table(action: "set-cell-text", ..., shape_index: 2, row: 1, column: 2, text: "Revenue")
```

Cells start empty and `row`/`column` are 1-based, so the header row is row `1`.

Size the table when you create it. Getting the height wrong and needing to change it later is
awkward, so estimate generously and confirm with a render.

### Structure

```
table(action: "insert-row", ..., before_row: 3)      # omit before_row to append
table(action: "delete-row", ..., row: 5)
table(action: "insert-column", ..., before_column: 2)
table(action: "merge-cells", ..., row: 1, column: 1, merge_to_row: 1, merge_to_column: 3)
```

Each of these returns the new `rowCount`/`columnCount`. Re-read it rather than assuming, and when
deleting several rows work from the highest index downwards so the earlier indexes stay valid.

### Formatting

```
table(action: "set-cell-fill", ..., row: 1, column: 1, red: 31, green: 78, blue: 121)
table(action: "set-cell-border", ..., row: 1, column: 1, border_type: "ppBorderBottom", weight: 2.0)
```

A shaded header row plus a single bottom border under it is usually enough. Heavy gridlines on
every cell make a slide look like a spreadsheet dump.

## Getting the data in

The natural workflow is: the assistant reads the numbers from wherever they live — a CSV, an API
response, a database query, or an Excel workbook via the sister
[Excel MCP Server](https://excelmcpserver.dev/) — and passes them straight into `add-chart` or
`set-cell-text`. No intermediate file, no copy and paste.

## Verify, always

Charts and tables are the two things most likely to look wrong despite every call succeeding —
axis labels colliding, a table taller than the slide, cell text wrapping to three lines.

```
export(action: "export-slide-to-image", session_id: ..., slide_index: 4, output_path: "C:\\preview\\slide4.png")
```

Look at the result before moving on. See
[Export slides to images](EXPORT-SLIDES-TO-IMAGES.md).

Common fixes:

| Symptom | Fix |
|---------|-----|
| Axis labels unreadable | Increase chart `width`/`height`, or use fewer categories |
| Pie chart unreadable | Use `"bar"` — pie only works below about six slices |
| Table text wrapping badly | Shorten cell text or widen the column set |
| Table overflowing the slide | Recreate it with a smaller `height`, or split across two slides |

## Related

- [Build a deck with AI](BUILD-A-DECK-WITH-AI.md)
- [Edit an existing deck](EDIT-AN-EXISTING-DECK.md)
- [Charts reference](https://powerpointmcpserver.dev/reference/charts/)
- [Tables reference](https://powerpointmcpserver.dev/reference/tables/)
