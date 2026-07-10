# Slides and Shapes

Reference for the `slide` tool (`add-blank`, `get-count`, `delete`, `duplicate`, `move-to`,
`set-background-color`, `get-background-color`, section management) and the `shape` tool
(`add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`, `add-connector`, `get-count`,
`delete`, `set-position`, `set-size`, plus the fill/line/rotation/flip/z-order/shadow/group/
name/alt-text/hyperlink formatting actions below).

## Slide Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `slide` | `add-blank` | `session_id` | Adds a **blank** slide at the end. No insert-at-index. |
| `slide` | `get-count` | `session_id` | Returns current slide count (`slideCount`). Call before/after mutations to confirm state. |
| `slide` | `delete` | `session_id`, `slide_index` (1-based) | Removes the slide; later slides shift down by one index. |
| `slide` | `duplicate` | `session_id`, `slide_index` | Inserts a copy of the slide immediately after the source. Returns the duplicate's new `slideIndex` and total `slideCount`. |
| `slide` | `move-to` | `session_id`, `slide_index`, `to_position` | Moves a slide to a new 1-based position, renumbering the rest. Returns the slide's new `slideIndex`. |
| `slide` | `set-background-color` | `session_id`, `slide_index`, `red`, `green`, `blue` | Sets a solid per-slide background color, overriding the slide master for that slide only. Returns `colorRgb` and `followsMasterBackground: false`. |
| `slide` | `get-background-color` | `session_id`, `slide_index` | Returns the slide's background `colorRgb` and `followsMasterBackground`. |
| `slide` | `add-section` | `session_id`, `section_index`, `section_name` (optional) | Adds a new section before `section_index` (pass `sectionCount + 1` to append). Returns the new `sectionIndex` and total `sectionCount`. |
| `slide` | `rename-section` | `session_id`, `section_index`, `section_name` | Renames an existing section. |
| `slide` | `delete-section` | `session_id`, `section_index`, `delete_slides` (optional, default `false`) | Deletes a section. If `delete_slides` is true, its slides are deleted too; otherwise they're kept and merged into a neighboring section. **PowerPoint disallows deleting section 1 unless `delete_slides` is true** — delete/reorder other sections first if you need to remove the first one's boundary without losing its slides. |
| `slide` | `get-section-count` | `session_id` | Returns the current number of sections (`sectionCount`, `0` if none exist). |
| `slide` | `get-section-name` | `session_id`, `section_index` | Returns a section's name (`sectionName`). |

Slides always append at the end via `add-blank` — there is no "insert blank at position N" action;
use `add-blank` then `move-to` if you need a blank slide inserted mid-deck. See `deck-builder.md`
for planning multi-slide order.

## Sections

Sections group contiguous ranges of slides for organizational purposes (visible in PowerPoint's
slide thumbnail panel) — they do not affect slide content or rendering. A slide's section
membership is purely positional (determined by where the slide sits between section boundaries),
so reordering slides with `move-to` can move them into a different section.

```
slide(action: "add-section", session_id: ..., section_index: 1, section_name: "Introduction")
slide(action: "add-section", session_id: ..., section_index: 2, section_name: "Deep Dive")
  → splits the deck into two sections at the current slide count boundary

slide(action: "get-section-count", session_id: ...) → sectionCount: 2
slide(action: "rename-section", session_id: ..., section_index: 2, section_name: "Details")
```

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

## Shape Formatting Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `shape` | `set-fill` | `session_id`, `slide_index`, `shape_index`, `red`, `green`, `blue` (each 0-255) | Sets a solid fill color. Returns `colorRgb`. |
| `shape` | `get-fill` | `session_id`, `slide_index`, `shape_index` | Returns the current fill color as `colorRgb`. |
| `shape` | `set-line` | `session_id`, `slide_index`, `shape_index`, plus optional `red`/`green`/`blue`, `weight`, `dash_style`, `visible` | All formatting params are optional and independently applied — pass only what you want to change. `red`/`green`/`blue` must be passed together to set the line color. `dash_style` is an `MsoLineDashStyle` name (see below). Returns the shape's full line state (`colorRgb`, `lineWeight`, `dashStyleName`, `visible`). |
| `shape` | `get-line` | `session_id`, `slide_index`, `shape_index` | Returns the current line color, weight, dash style, and visibility. |
| `shape` | `set-rotation` | `session_id`, `slide_index`, `shape_index`, `degrees` | Sets rotation in degrees clockwise from upright. Returns `rotation`. |
| `shape` | `get-rotation` | `session_id`, `slide_index`, `shape_index` | Returns the current rotation in degrees. |
| `shape` | `flip` | `session_id`, `slide_index`, `shape_index`, `direction` (`horizontal` or `vertical`) | Flips the shape in place. Returns `flipDirection`. |
| `shape` | `set-z-order` | `session_id`, `slide_index`, `shape_index`, `z_order_command` | Moves the shape's stacking position. `z_order_command` is one of `bring-to-front`, `send-to-back`, `bring-forward`, `send-backward`. Returns `zOrderCommand`. |
| `shape` | `set-shadow` | `session_id`, `slide_index`, `shape_index`, `visible` | Turns the shape's default drop shadow on/off. Returns `visible`. |
| `shape` | `get-shadow` | `session_id`, `slide_index`, `shape_index` | Returns whether the shadow is visible. |
| `shape` | `group` | `session_id`, `slide_index`, `shape_indexes` (JSON array of 1-based indices, at least 2) | Groups multiple shapes into one. Returns the new total `shapeCount` on the slide — **not** the grouped shape's index (see NoPIA note below). |
| `shape` | `ungroup` | `session_id`, `slide_index`, `shape_index` | Splits a group back into its member shapes. Returns `ungroupedShapeCount` (members produced) and the new total `shapeCount`. |
| `shape` | `set-name` | `session_id`, `slide_index`, `shape_index`, `name` | Sets the shape's name (as shown in PowerPoint's Selection Pane). Returns `name`. |
| `shape` | `get-name` | `session_id`, `slide_index`, `shape_index` | Returns the shape's current name. |
| `shape` | `set-alt-text` | `session_id`, `slide_index`, `shape_index`, `alt_text` | Sets the shape's alternative text (accessibility description). Returns `altText`. |
| `shape` | `get-alt-text` | `session_id`, `slide_index`, `shape_index` | Returns the shape's current alternative text. |
| `shape` | `set-hyperlink` | `session_id`, `slide_index`, `shape_index`, `address`, `screen_tip` (optional) | Sets the shape's mouse-click hyperlink to `address` (URL or file path). `screen_tip` sets hover tooltip text. Returns `hasHyperlink`, `hyperlinkAddress`, `hyperlinkScreenTip`. |
| `shape` | `get-hyperlink` | `session_id`, `slide_index`, `shape_index` | Returns `hasHyperlink` and, if present, `hyperlinkAddress`/`hyperlinkScreenTip`. |
| `shape` | `remove-hyperlink` | `session_id`, `slide_index`, `shape_index` | Removes the shape's mouse-click hyperlink, if any (no-op if none is set). Returns `hasHyperlink: false`. |

**Finding a just-grouped shape's index**: `group` does not return the new group shape's own
`shapeIndex` — reading `.Index` off a freshly-created COM group object is unreliable in this
codebase's NoPIA late-binding setup. If the shapes you grouped were the **last** shapes added to
the slide (highest indices, nothing added after them), the resulting group occupies the new,
smaller `shapeCount` as its index (since grouping N shapes always removes N-1 from the slide's
shape list). Otherwise, call `shape(action: "get-count", ...)` before and after grouping and
inspect via a follow-up read if you need to confirm which index now holds the group.

### Hyperlinks

`set-hyperlink`/`get-hyperlink`/`remove-hyperlink` manage a shape's **mouse-click** action —
clicking the shape at presentation time navigates to `address` (an absolute URL like
`"https://example.com"`, or a local file path). There is no separate mouse-hover hyperlink action,
and no text-run-level hyperlink (a whole-shape hyperlink is the only granularity this tool
surface exposes) — to make specific words within a text box clickable, put that text in its own
shape.

`address` is normalized by PowerPoint itself (e.g. `"https://example.com"` round-trips as
`"https://example.com/"` with a trailing slash) — compare against the returned `hyperlinkAddress`
rather than assuming an exact byte-for-byte match of what you passed in.

```
shape(action: "set-hyperlink", session_id: ..., slide_index: ..., shape_index: ...,
  address: "https://example.com", screen_tip: "Visit our site")
shape(action: "get-hyperlink", session_id: ..., slide_index: ..., shape_index: ...)
  → hasHyperlink: true, hyperlinkAddress: "https://example.com/", hyperlinkScreenTip: "Visit our site"
shape(action: "remove-hyperlink", session_id: ..., slide_index: ..., shape_index: ...)
  → hasHyperlink: false
```

### Dash Styles (`dash_style` for `set-line`)

Must match a real `MsoLineDashStyle` enum member name exactly: `msoLineSolid`,
`msoLineSquareDot`, `msoLineRoundDot`, `msoLineDash`, `msoLineDashDot`, `msoLineDashDotDot`,
`msoLineLongDash`, `msoLineLongDashDot`.

### Z-Order Commands (`z_order_command` for `set-z-order`)

`bring-to-front`, `send-to-back`, `bring-forward`, `send-backward`. (PowerPoint's Word-only
z-order members — bring/send relative to text — are intentionally not exposed here.)


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
