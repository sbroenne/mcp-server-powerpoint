# Tables

Reference for `table(action: "add-table", ...)`, `table(action: "set-cell-text", ...)`,
`table(action: "get-cell-text", ...)`.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `table` | `add-table` | `session_id`, `slide_index`, `rows`, `columns`, `left`, `top`, `width`, `height` | Creates a table shape with the given dimensions; cells start empty. Returns `shapeIndex`. |
| `table` | `set-cell-text` | `session_id`, `slide_index`, `shape_index`, `row`, `column`, `text` | `row`/`column` are 1-based. |
| `table` | `get-cell-text` | `session_id`, `slide_index`, `shape_index`, `row`, `column` | Reads a single cell's text (`cellText`). |

## Fixed Dimensions

`rows` and `columns` are set at creation time by `add-table` — there is no "add row"/"add column"
action in this surface. Decide the final table shape (including a header row, if any) before
calling `add-table`. If you need a different size later, delete the table shape (`shape(action:
"delete", ...)`) and re-create it with `add-table` at the corrected dimensions — cell content is
not preserved automatically, so re-populate all cells after resizing this way.

## Populating a Table

Header row is just row 1 — set it like any other row, typically with `set-bold`/`set-font-size`
applied via the TextFrame tools once the cell text is set (tables share the shape/text-frame
model, so `textframe(action: "set-bold"/"set-font-size"/"set-font-color", ...)` also work against
a table's `shape_index` if you want uniform emphasis across the whole table; per-cell formatting
beyond text content is not exposed by this tool surface).

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
