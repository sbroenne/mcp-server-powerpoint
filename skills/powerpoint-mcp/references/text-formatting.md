# Text Formatting: TextFrame Tools

Reference for `set_text`, `get_text`, `set_font_size`, `set_bold`, `set_font_color` — all operate
on a shape's text frame.

## Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `set_text` | `sessionId`, `slideIndex`, `shapeIndex`, `text` | Replaces the shape's entire text content. |
| `get_text` | `sessionId`, `slideIndex`, `shapeIndex` | Reads current text — use before editing to avoid clobbering unrelated content. |
| `set_font_size` | `sessionId`, `slideIndex`, `shapeIndex`, `fontSize` (points) | Applies to the shape's **entire** text range, not a substring. |
| `set_bold` | `sessionId`, `slideIndex`, `shapeIndex`, `bold` (bool) | Applies to the entire text range. |
| `set_font_color` | `sessionId`, `slideIndex`, `shapeIndex`, `red`, `green`, `blue` (each 0-255) | RGB triplet, applies to the entire text range. |

## Whole-Range Formatting Only

These formatting tools apply to a shape's **entire** text frame — there is no API here for
formatting a substring or a specific run of characters within one text box. If a slide needs
mixed formatting (e.g., a bold label next to plain description text), use **separate text boxes**
positioned next to each other rather than trying to mix runs inside one shape:

```
CORRECT — two text boxes for mixed emphasis
add_text_box(sessionId, slideIndex, left=50, top=100, width=150, height=30, text="Revenue:")
set_bold(sessionId, slideIndex, <label shapeIndex>, bold=true)
add_text_box(sessionId, slideIndex, left=210, top=100, width=300, height=30, text="$2.4M, up 12%")
```

## Bullet Lists via Line Breaks

There is no dedicated "bullet list" tool. Build a bulleted look by embedding line breaks (`\n`)
in the `text` parameter of `add_text_box` / `set_text`, with a leading dash or similar marker per
line:

```
set_text(sessionId, slideIndex, shapeIndex,
  text="- Revenue grew 24% year over year\n- APAC now the fastest-growing region\n- Retention held steady at 91%")
```

Keep each line short — this is plain text, not PowerPoint's native outline/bullet formatting, so
there's no auto-indent or bullet glyph; a leading `-` or `•` character is the simplest visual cue.

## Font Size Guidance

| Element | Suggested `fontSize` |
|---------|----------------------|
| Slide title | 28-40 |
| Subtitle | 18-24 |
| Body text | 14-18 |
| Table cell text | see `tables.md` |
| Captions / footnotes | 10-12 |

Keep to 2-3 distinct font sizes per slide; titles are typically bold, body text typically is not.

## Color Values

`set_font_color` takes three separate `byte` parameters (`red`, `green`, `blue`), each 0-255 —
**not** a hex string. Convert a hex color to decimal triplets before calling:

```
Hex "4472C4" → red=68, green=114, blue=196
set_font_color(sessionId, slideIndex, shapeIndex, red=68, green=114, blue=196)
```

## Read Before You Overwrite

`set_text` replaces the whole text frame content. If you only need to append or tweak part of an
existing shape's text, call `get_text` first, compose the full new string yourself, and pass the
complete result to `set_text` — there is no append/insert operation.

```
get_text(sessionId, slideIndex, shapeIndex) → "Q3 Results"
set_text(sessionId, slideIndex, shapeIndex, text="Q3 Results (Final)")
```

## Verify Text Fit

After setting text and font size, export the slide (see `export-and-verify.md`) to confirm the
text fits inside the shape's `width`/`height` without visually overflowing — there is no
auto-shrink-to-fit in this tool surface. If text overflows: shorten the text, reduce
`set_font_size`, or grow the shape with `set_shape_size`.
