# Slides and Shapes

Reference for `add_slide`, `get_slide_count`, `delete_slide`, and the `Shape` tools:
`add_rectangle`, `add_text_box`, `get_shape_count`, `delete_shape`, `set_shape_position`,
`set_shape_size`.

## Slide Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `add_slide` | `sessionId` | Adds a **blank** slide at the end. No insert-at-index. |
| `get_slide_count` | `sessionId` | Returns current slide count. Call before/after mutations to confirm state. |
| `delete_slide` | `sessionId`, `slideIndex` (1-based) | Removes the slide; later slides shift down by one index. |

Slides always append at the end — there is no "insert at position N" tool. See `deck-builder.md`
for planning multi-slide order.

## Shape Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `add_rectangle` | `sessionId`, `slideIndex`, `left`, `top`, `width`, `height` | Plain rectangle, no fill/line color parameters — style comes from PowerPoint's theme default. |
| `add_text_box` | `sessionId`, `slideIndex`, `left`, `top`, `width`, `height`, `text` | Creates the text box AND sets its initial text in one call. |
| `get_shape_count` | `sessionId`, `slideIndex` | Number of shapes currently on the slide. |
| `delete_shape` | `sessionId`, `slideIndex`, `shapeIndex` (1-based) | Removes one shape; later shapes on that slide shift down by one index. |
| `set_shape_position` | `sessionId`, `slideIndex`, `shapeIndex`, `left`, `top` | Moves an existing shape. |
| `set_shape_size` | `sessionId`, `slideIndex`, `shapeIndex`, `width`, `height` | Resizes an existing shape. |

All position/size values are **points** (see `deck-builder.md` for the 960×540pt 16:9 reference).

## Shape Indexing Within a Slide

`shapeIndex` is 1-based and reflects the **order shapes were added to that slide** (and any
built-in placeholders from the applied layout, if present). After adding several shapes, use
`get_shape_count` to confirm the current total before referencing an index you didn't just
create yourself — don't assume index 1 is always the title.

```
add_text_box(sessionId, 1, ..., text="Title")   → shapeIndex 1 (assuming a blank slide)
add_rectangle(sessionId, 1, ...)                 → shapeIndex 2
add_text_box(sessionId, 1, ..., text="Body")     → shapeIndex 3
get_shape_count(sessionId, 1) → 3                → confirms the count before further edits
```

## Building a Text Box + Table/Chart Combo Slide

Tables and charts are added by their own domain tools (`add_table`, `add_chart` — see
`tables.md` and `charts.md`) but they are shapes on the slide like any other, and share the same
`shapeIndex` numbering with rectangles and text boxes added on that slide. Track the returned
`shapeIndex` from each add call (or re-check with `get_shape_count`) so subsequent
position/size/format calls target the right shape.

## Repositioning and Resizing

Use `set_shape_position` / `set_shape_size` for targeted layout fixes instead of deleting and
re-adding a shape:

```
CORRECT — nudge a shape that overlaps another after visual verification
set_shape_position(sessionId, slideIndex, shapeIndex, left=500, top=120)

AVOID — delete and recreate to move a shape
delete_shape(sessionId, slideIndex, shapeIndex)
add_rectangle(sessionId, slideIndex, left=500, top=120, width=..., height=...)
```

Deleting and recreating loses the shape's text content and any formatting already applied to it —
prefer targeted `set_shape_position`/`set_shape_size` (see `anti-patterns.md`).
