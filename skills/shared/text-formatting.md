# Text Formatting: TextFrame Tools

Reference for the `textframe` tool's actions — `set-text`, `get-text`, `set-font-size`,
`set-bold`, `set-font-color` — all operate on a shape's text frame.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `textframe` | `set-text` | `session_id`, `slide_index`, `shape_index`, `text` | Replaces the shape's entire text content. |
| `textframe` | `get-text` | `session_id`, `slide_index`, `shape_index` | Reads current text (`text`) — use before editing to avoid clobbering unrelated content. |
| `textframe` | `set-font-size` | `session_id`, `slide_index`, `shape_index`, `font_size` (points) | Applies to the shape's **entire** text range, not a substring. |
| `textframe` | `set-bold` | `session_id`, `slide_index`, `shape_index`, `bold` (bool) | Applies to the entire text range. |
| `textframe` | `set-font-color` | `session_id`, `slide_index`, `shape_index`, `red`, `green`, `blue` (each 0-255) | RGB triplet, applies to the entire text range. |

## Whole-Range Formatting Only

These formatting actions apply to a shape's **entire** text frame — there is no API here for
formatting a substring or a specific run of characters within one text box. If a slide needs
mixed formatting (e.g., a bold label next to plain description text), use **separate text boxes**
positioned next to each other rather than trying to mix runs inside one shape:

```
CORRECT — two text boxes for mixed emphasis
shape(action: "add-text-box", session_id: ..., slide_index: ..., left: 50, top: 100, width: 150, height: 30, text: "Revenue:")
textframe(action: "set-bold", session_id: ..., slide_index: ..., shape_index: <label shapeIndex>, bold: true)
shape(action: "add-text-box", session_id: ..., slide_index: ..., left: 210, top: 100, width: 300, height: 30, text: "$2.4M, up 12%")
```

## Bullet Lists via Line Breaks

There is no dedicated "bullet list" action. Build a bulleted look by embedding line breaks (`\n`)
in the `text` parameter of `shape(action: "add-text-box", ...)` / `textframe(action: "set-text",
...)`, with a leading dash or similar marker per line:

```
textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ...,
  text: "- Revenue grew 24% year over year\n- APAC now the fastest-growing region\n- Retention held steady at 91%")
```

Keep each line short — this is plain text, not PowerPoint's native outline/bullet formatting, so
there's no auto-indent or bullet glyph; a leading `-` or `•` character is the simplest visual cue.

## Font Size Guidance

| Element | Suggested `font_size` |
|---------|----------------------|
| Slide title | 28-40 |
| Subtitle | 18-24 |
| Body text | 14-18 |
| Table cell text | see `tables.md` |
| Captions / footnotes | 10-12 |

Keep to 2-3 distinct font sizes per slide; titles are typically bold, body text typically is not.

## Color Values

`set-font-color` takes three separate `byte` parameters (`red`, `green`, `blue`), each 0-255 —
**not** a hex string. Convert a hex color to decimal triplets before calling:

```
Hex "4472C4" → red=68, green=114, blue=196
textframe(action: "set-font-color", session_id: ..., slide_index: ..., shape_index: ..., red: 68, green: 114, blue: 196)
```

## Read Before You Overwrite

`set-text` replaces the whole text frame content. If you only need to append or tweak part of an
existing shape's text, call `get-text` first, compose the full new string yourself, and pass the
complete result to `set-text` — there is no append/insert operation.

```
textframe(action: "get-text", session_id: ..., slide_index: ..., shape_index: ...) → "Q3 Results"
textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ..., text: "Q3 Results (Final)")
```

## Verify Text Fit

After setting text and font size, export the slide (see `export-and-verify.md`) to confirm the
text fits inside the shape's `width`/`height` without visually overflowing — there is no
auto-shrink-to-fit in this tool surface. If text overflows: shorten the text, reduce
`font_size`, or grow the shape with `shape(action: "set-size", ...)`.
