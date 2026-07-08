# Slides and Shapes

Reference for the `slide` tool (`add-blank`, `get-count`, `delete`) and the `shape` tool
(`add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`, `add-connector`, `get-count`,
`delete`, `set-position`, `set-size`).

## Slide Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `slide` | `add-blank` | `session_id` | Adds a **blank** slide at the end. No insert-at-index. |
| `slide` | `get-count` | `session_id` | Returns current slide count (`slideCount`). Call before/after mutations to confirm state. |
| `slide` | `delete` | `session_id`, `slide_index` (1-based) | Removes the slide; later slides shift down by one index. |

Slides always append at the end — there is no "insert at position N" action. See
`deck-builder.md` for planning multi-slide order.

## Shape Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `shape` | `add-rectangle` | `session_id`, `slide_index`, `left`, `top`, `width`, `height` | Plain rectangle, no fill/line color parameters — style comes from PowerPoint's theme default. Returns `shapeIndex`. |
| `shape` | `add-text-box` | `session_id`, `slide_index`, `left`, `top`, `width`, `height`, `text` | Creates the text box AND sets its initial text in one call. Returns `shapeIndex`. |
| `shape` | `add-auto-shape` | `session_id`, `slide_index`, `shape_type`, `left`, `top`, `width`, `height` | Adds any non-rectangle built-in shape (oval, diamond, arrow, star bracket, etc.) by its `MsoAutoShapeType` name. Returns `shapeIndex` and echoes `shapeTypeName`. See "Auto Shape Types" below for the supported name list. |
| `shape` | `add-line` | `session_id`, `slide_index`, `begin_x`, `begin_y`, `end_x`, `end_y` | Straight line between two points. Returns `shapeIndex` and echoes `beginX`/`beginY`/`endX`/`endY`. |
| `shape` | `add-connector` | `session_id`, `slide_index`, `connector_type`, `begin_x`, `begin_y`, `end_x`, `end_y` | Adds a connector shape (`msoConnectorStraight`, `msoConnectorElbow`, or `msoConnectorCurve`) between two points. Free-floating — not glued to other shapes. Returns `shapeIndex` and echoes `connectorTypeName`. |
| `shape` | `get-count` | `session_id`, `slide_index` | Number of shapes currently on the slide (`shapeCount`). |
| `shape` | `delete` | `session_id`, `slide_index`, `shape_index` (1-based) | Removes one shape; later shapes on that slide shift down by one index. |
| `shape` | `set-position` | `session_id`, `slide_index`, `shape_index`, `left`, `top` | Moves an existing shape. |
| `shape` | `set-size` | `session_id`, `slide_index`, `shape_index`, `width`, `height` | Resizes an existing shape. |

All position/size values are **points** (see `deck-builder.md` for the 960×540pt 16:9 reference).

## Auto Shape Types

`shape_type` for `add-auto-shape` must match a real `MsoAutoShapeType` enum member name exactly
(case-sensitive, `mso`-prefixed PascalCase) — this is a curated subset (not the full Office enum):

| Category | `shape_type` values |
|----------|--------------------|
| Basic | `msoShapeRectangle`, `msoShapeRoundedRectangle`, `msoShapeOval`, `msoShapeDiamond`, `msoShapeParallelogram`, `msoShapeTrapezoid`, `msoShapeIsoscelesTriangle`, `msoShapeRightTriangle`, `msoShapeHexagon`, `msoShapeOctagon`, `msoShapeRegularPentagon`, `msoShapeCross` |
| Arrows | `msoShapeRightArrow`, `msoShapeLeftArrow`, `msoShapeUpArrow`, `msoShapeDownArrow`, `msoShapeLeftRightArrow`, `msoShapeUpDownArrow` |
| Brackets/braces | `msoShapeLeftBracket`, `msoShapeRightBracket`, `msoShapeLeftBrace`, `msoShapeRightBrace` |
| Decorative/misc | `msoShapeCan`, `msoShapeCube`, `msoShapeBevel`, `msoShapeFoldedCorner`, `msoShapeSmileyFace`, `msoShapeDonut`, `msoShapeNoSymbol`, `msoShapeBlockArc`, `msoShapeHeart`, `msoShapeLightningBolt`, `msoShapeSun`, `msoShapeMoon`, `msoShapeArc`, `msoShapePlaque` |

Passing an unrecognized string returns `success: false` — double-check spelling rather than
guessing variants (e.g. star/callout shapes are not in this curated set).

For lines and connectors, `connector_type` (add-connector only) must be one of
`msoConnectorStraight`, `msoConnectorElbow`, or `msoConnectorCurve`.

## Shape Indexing Within a Slide

`shape_index` is 1-based and reflects the **order shapes were added to that slide** (and any
built-in placeholders from the applied layout, if present). After adding several shapes, use
`shape(action: "get-count", ...)` to confirm the current total before referencing an index you
didn't just create yourself — don't assume index 1 is always the title.

```
shape(action: "add-text-box", session_id: ..., slide_index: 1, ..., text: "Title")   → shapeIndex 1 (assuming a blank slide)
shape(action: "add-rectangle", session_id: ..., slide_index: 1, ...)                  → shapeIndex 2
shape(action: "add-text-box", session_id: ..., slide_index: 1, ..., text: "Body")     → shapeIndex 3
shape(action: "get-count", session_id: ..., slide_index: 1) → 3                        → confirms the count before further edits
```

## Building a Text Box + Table/Chart Combo Slide

Tables and charts are added by their own domain tools (`table(action: "add-table", ...)`,
`chart(action: "add-chart", ...)` — see `tables.md` and `charts.md`) but they are shapes on the
slide like any other, and share the same `shape_index` numbering with rectangles and text boxes
added on that slide. Track the returned `shapeIndex` from each add call (or re-check with
`shape(action: "get-count", ...)`) so subsequent position/size/format calls target the right
shape.

## Repositioning and Resizing

Use `shape(action: "set-position", ...)` / `shape(action: "set-size", ...)` for targeted layout
fixes instead of deleting and re-adding a shape:

```
CORRECT — nudge a shape that overlaps another after visual verification
shape(action: "set-position", session_id: ..., slide_index: ..., shape_index: ..., left: 500, top: 120)

AVOID — delete and recreate to move a shape
shape(action: "delete", session_id: ..., slide_index: ..., shape_index: ...)
shape(action: "add-rectangle", session_id: ..., slide_index: ..., left: 500, top: 120, width: ..., height: ...)
```

Deleting and recreating loses the shape's text content and any formatting already applied to it —
prefer targeted `set-position`/`set-size` (see `anti-patterns.md`).
