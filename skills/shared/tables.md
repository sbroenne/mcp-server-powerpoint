# Tables

Reference for `add_table`, `set_cell_text`, `get_cell_text`.

## Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `add_table` | `sessionId`, `slideIndex`, `rows`, `columns`, `left`, `top`, `width`, `height` | Creates a table shape with the given dimensions; cells start empty. |
| `set_cell_text` | `sessionId`, `slideIndex`, `shapeIndex`, `row`, `column`, `text` | `row`/`column` are 1-based. |
| `get_cell_text` | `sessionId`, `slideIndex`, `shapeIndex`, `row`, `column` | Reads a single cell's text. |

## Fixed Dimensions

`rows` and `columns` are set at creation time by `add_table` — there is no "add row"/"add column"
tool in this surface. Decide the final table shape (including a header row, if any) before
calling `add_table`. If you need a different size later, delete the table shape
(`delete_shape`) and re-create it with `add_table` at the corrected dimensions — cell content
is not preserved automatically, so re-populate all cells after resizing this way.

## Populating a Table

Header row is just row 1 — set it like any other row, typically with `set_bold`/`set_font_size`
applied via the TextFrame tools once the cell text is set (tables share the shape/text-frame
model, so `set_bold`/`set_font_size`/`set_font_color` also work against a table's `shapeIndex`
if you want uniform emphasis across the whole table; per-cell formatting beyond text content is
not exposed by this tool surface).

```
add_table(sessionId, slideIndex, rows=4, columns=3, left=60, top=120, width=500, height=250)
  → shapeIndex (from get_shape_count, or track return value)

# Header row
set_cell_text(sessionId, slideIndex, shapeIndex, row=1, column=1, text="Region")
set_cell_text(sessionId, slideIndex, shapeIndex, row=1, column=2, text="Revenue")
set_cell_text(sessionId, slideIndex, shapeIndex, row=1, column=3, text="Growth")

# Data rows
set_cell_text(sessionId, slideIndex, shapeIndex, row=2, column=1, text="APAC")
set_cell_text(sessionId, slideIndex, shapeIndex, row=2, column=2, text="$2.4M")
set_cell_text(sessionId, slideIndex, shapeIndex, row=2, column=3, text="+24%")
# ... continue for remaining rows ...
```

## Targeted Cell Updates

Update only the cells that changed rather than re-populating the whole table:

```
CORRECT — fix one cell
set_cell_text(sessionId, slideIndex, shapeIndex, row=3, column=2, text="$1.8M")

AVOID — delete and rebuild the whole table for a one-cell change
delete_shape(sessionId, slideIndex, shapeIndex)
add_table(...)
# ... re-populate every cell ...
```

## Sizing Guidance

- Keep cell text short — table cells have no dedicated font-size default beyond the shape/text
  formatting tools, and long strings will visually overflow at typical `width`/`rows` ratios.
  Suggested body text: 12-14pt when set via `set_font_size`.
- Reserve enough `height` for the number of `rows` — a rough guide is `height ≈ rows × 30-40pt`
  for readable single-line rows.
- Verify with `export_slide_to_image` after populating — text overflow inside table cells is a
  common visual issue that only shows up in the rendered image.

## Reading Existing Tables

Use `get_cell_text` to inspect specific cells before overwriting — there is no "get whole table"
call, so read the cells you need individually (e.g., every cell in a row you're about to edit) if
you need to preserve unrelated content.
