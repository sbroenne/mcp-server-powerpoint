# Tables

Reference for `table` actions: creating tables, reading/writing cell text, inserting/deleting rows
and columns, formatting cell fill and borders, and merging cells.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `table` | `add-table` | `session_id`, `slide_index`, `rows`, `columns`, `left`, `top`, `width`, `height` | Creates a table shape with the given dimensions; cells start empty. Returns `shapeIndex`. |
| `table` | `set-cell-text` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `text` | `row`/`column` are 1-based. |
| `table` | `get-cell-text` | `session_id`, `slide_index`, `shape_index`, `row`, `column` | Reads a single cell's text (`cellText`). |
| `table` | `insert-row` | `session_id`, `slide_index`, `shape_index`, `before_row` (optional) | Inserts a new row before `before_row`; omit to append as the last row. Returns new `rowCount`. |
| `table` | `delete-row` | `session_id`, `slide_index`, `shape_index`, `row` | Deletes the row at `row`. Returns new `rowCount`. |
| `table` | `insert-column` | `session_id`, `slide_index`, `shape_index`, `before_column` (optional) | Inserts a new column before `before_column`; omit to append as the last column. Returns new `columnCount`. |
| `table` | `delete-column` | `session_id`, `slide_index`, `shape_index`, `column` | Deletes the column at `column`. Returns new `columnCount`. |
| `table` | `set-cell-fill` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `red`, `green`, `blue` | Sets a cell's solid fill color (0-255 per channel). Returns `colorRgb`. |
| `table` | `get-cell-fill` | `session_id`, `slide_index`, `shape_index`, `row`, `column` | Reads a cell's solid fill color (`colorRgb`). |
| `table` | `set-cell-border` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `border_type`, `red`/`green`/`blue` (optional), `weight` (optional), `dash_style` (optional), `visible` (optional) | Sets one border of a cell. All formatting params optional except `border_type`; omit a param to leave it unchanged. |
| `table` | `get-cell-border` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `border_type` | Reads a cell border's color, weight, dash style, and visibility. |
| `table` | `merge-cells` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `merge_to_row`, `merge_to_column` | Merges two adjacent cells into one. Returns new `rowCount`/`columnCount` (unchanged — grid dimensions stay the same after a merge). |

## `border_type` values

`ppBorderTop`, `ppBorderLeft`, `ppBorderBottom`, `ppBorderRight`, `ppBorderDiagonalDown`,
`ppBorderDiagonalUp` (`PpBorderType` enum member names).

## `dash_style` values

Same curated `MsoLineDashStyle` subset as `shape(action: "set-line", ...)`: `msoLineSolid`,
`msoLineSquareDot`, `msoLineRoundDot`, `msoLineDash`, `msoLineDashDot`, `msoLineDashDotDot`,
`msoLineLongDash`, `msoLineLongDashDot`, `msoLineLongDashDotDot`.

## Editing Table Structure

Rows and columns no longer need to be fixed at creation time — use `insert-row`/`delete-row` and
`insert-column`/`delete-column` to adjust an existing table's shape. Existing cell content in
unaffected rows/columns is preserved automatically; there is no need to delete and re-create the
table shape for a structural change.

```
table(action: "insert-row", session_id: ..., slide_index: ..., shape_index: ..., before_row: 2)
  → rowCount

table(action: "delete-column", session_id: ..., slide_index: ..., shape_index: ..., column: 4)
  → columnCount
```

## Populating a Table

Header row is just row 1 — set it like any other row, typically with `set-bold`/`set-font-size`
applied via the TextFrame tools once the cell text is set (tables share the shape/text-frame
model, so `textframe(action: "set-bold"/"set-font-size"/"set-font-color", ...)` also work against
a table's `shape_index`). Per-cell fill color and borders can be set directly with
`set-cell-fill`/`set-cell-border` — e.g. shading the header row or highlighting a specific data
cell — without affecting the rest of the table.

```
table(action: "add-table", session_id: ..., slide_index: ..., rows: 4, columns: 3, left: 60, top: 120, width: 500, height: 250)
  → shapeIndex

# Header row
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 1, column: 1, text: "Region")
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 1, column: 2, text: "Revenue")
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 1, column: 3, text: "Growth")

# Data rows
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 2, column: 1, text: "APAC")
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 2, column: 2, text: "$2.4M")
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 2, column: 3, text: "+24%")
# ... continue for remaining rows ...
```

## Targeted Cell Updates

Update only the cells that changed rather than re-populating the whole table:

```
CORRECT — fix one cell
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 3, column: 2, text: "$1.8M")

AVOID — delete and rebuild the whole table for a one-cell change
shape(action: "delete", session_id: ..., slide_index: ..., shape_index: ...)
table(action: "add-table", ...)
# ... re-populate every cell ...
```

## Sizing Guidance

- Keep cell text short — table cells have no dedicated font-size default beyond the shape/text
  formatting tools, and long strings will visually overflow at typical `width`/`rows` ratios.
  Suggested body text: 12-14pt when set via `textframe(action: "set-font-size", ...)`.
- Reserve enough `height` for the number of `rows` — a rough guide is `height ≈ rows × 30-40pt`
  for readable single-line rows.
- Verify with `export(action: "export-slide-to-image", ...)` after populating — text overflow
  inside table cells is a common visual issue that only shows up in the rendered image.

## Reading Existing Tables

Use `get-cell-text` to inspect specific cells before overwriting — there is no "get whole table"
action, so read the cells you need individually (e.g., every cell in a row you're about to edit)
if you need to preserve unrelated content.
