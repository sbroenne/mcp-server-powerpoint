# Text Formatting: TextFrame Tools

Reference for the `textframe` tool's actions — `set-text`, `get-text`, `set-font-size`,
`set-bold`, `set-font-color`, `set-italic`/`get-italic`, `set-underline`/`get-underline`,
`set-font-name`/`get-font-name`, `set-alignment`/`get-alignment`, `set-bullet`/`get-bullet`,
`set-autosize`/`get-autosize` — all operate on a shape's text frame.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `textframe` | `set-text` | `session_id`, `slide_index`, `shape_index`, `text` | Replaces the shape's entire text content. |
| `textframe` | `get-text` | `session_id`, `slide_index`, `shape_index` | Reads current text (`text`) — use before editing to avoid clobbering unrelated content. |
| `textframe` | `set-font-size` | `session_id`, `slide_index`, `shape_index`, `font_size` (points) | Applies to the shape's **entire** text range, not a substring. |
| `textframe` | `set-bold` | `session_id`, `slide_index`, `shape_index`, `bold` (bool) | Applies to the entire text range. |
| `textframe` | `set-font-color` | `session_id`, `slide_index`, `shape_index`, `red`, `green`, `blue` (each 0-255) | RGB triplet, applies to the entire text range. |
| `textframe` | `set-italic` | `session_id`, `slide_index`, `shape_index`, `italic` (bool) | Applies to the entire text range. |
| `textframe` | `get-italic` | `session_id`, `slide_index`, `shape_index` | Returns `italic`. |
| `textframe` | `set-underline` | `session_id`, `slide_index`, `shape_index`, `underline` (bool) | Applies to the entire text range. |
| `textframe` | `get-underline` | `session_id`, `slide_index`, `shape_index` | Returns `underline`. |
| `textframe` | `set-font-name` | `session_id`, `slide_index`, `shape_index`, `font_name` | Sets the typeface (e.g. `"Calibri"`, `"Georgia"`). No validation against installed fonts — an unrecognized name silently falls back to a substitute font in PowerPoint. |
| `textframe` | `get-font-name` | `session_id`, `slide_index`, `shape_index` | Returns `fontName`. |
| `textframe` | `set-alignment` | `session_id`, `slide_index`, `shape_index`, `alignment` | Sets paragraph alignment. `alignment` is a `PpParagraphAlignment` name (see below). |
| `textframe` | `get-alignment` | `session_id`, `slide_index`, `shape_index` | Returns `alignment`. Fails if paragraphs within the text range have mixed alignment. |
| `textframe` | `set-bullet` | `session_id`, `slide_index`, `shape_index`, `enabled` (bool), optional `character` (single character) | Turns bullets on/off for every paragraph in the text range. When enabling, `character` sets the bullet glyph (e.g. `"-"`, `"•"`); omit to keep the theme's default bullet. |
| `textframe` | `get-bullet` | `session_id`, `slide_index`, `shape_index` | Returns `bulletEnabled` and `bulletCharacter` (null when bullets are off). |
| `textframe` | `set-autosize` | `session_id`, `slide_index`, `shape_index`, `auto_size` | Sets the text frame's auto-fit behavior. `auto_size` is a `PpAutoSize` name (see below). |
| `textframe` | `get-autosize` | `session_id`, `slide_index`, `shape_index` | Returns `autoSize`. Fails if the value is mixed across multiple shapes. |

## Paragraph Alignment Names

`alignment` for `set-alignment` must match a real `PpParagraphAlignment` enum member name exactly:
`ppAlignLeft`, `ppAlignCenter`, `ppAlignRight`, `ppAlignJustify`, `ppAlignDistribute`,
`ppAlignThaiDistribute`, `ppAlignJustifyLow`. Passing an unrecognized string returns
`success: false`.

## Auto-Fit / Auto-Size

`auto_size` for `set-autosize` must match a real `PpAutoSize` enum member name exactly:

| `auto_size` | Behavior |
|-------------|----------|
| `ppAutoSizeNone` | No auto-fit — text can overflow the shape's bounds (the default for most shapes). |
| `ppAutoSizeShapeToFitText` | The **shape grows/shrinks** to fit its text; text stays at its set font size. |
| `ppAutoSizeTextToFitShape` | The **text shrinks** (font scales down) to fit inside a fixed-size shape — PowerPoint's "Shrink text on overflow". |

Passing an unrecognized string returns `success: false`. Set this **after** `set-text` and any
font-size changes — `ppAutoSizeTextToFitShape` computes the shrink factor from whatever text is in
the frame at the moment PowerPoint next reflows it.

```
textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ..., text: "A long paragraph that might overflow the box...")
textframe(action: "set-autosize", session_id: ..., slide_index: ..., shape_index: ..., auto_size: "ppAutoSizeTextToFitShape")
```

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

## Bullet Lists

Use `textframe(action: "set-bullet", ...)` for native PowerPoint bullets — it applies PowerPoint's
own bullet glyph and per-paragraph indent, unlike a plain leading `-`/`•` character embedded in
text. Put each bullet item on its own line (`\n`-separated) in the `text` parameter first, then
turn bullets on for the whole text range:

```
textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ...,
  text: "Revenue grew 24% year over year\nAPAC now the fastest-growing region\nRetention held steady at 91%")
textframe(action: "set-bullet", session_id: ..., slide_index: ..., shape_index: ..., enabled: true)
```

If a specific bullet glyph is required (e.g. a dash instead of the theme's default), pass
`character`: `textframe(action: "set-bullet", ..., enabled: true, character: "-")`.

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
text fits inside the shape's `width`/`height` without visually overflowing. Use
`set-autosize` (see "Auto-Fit / Auto-Size" above) to have PowerPoint handle overflow
automatically instead of manually tuning size — `ppAutoSizeTextToFitShape` shrinks the font to
fit a fixed box, `ppAutoSizeShapeToFitText` grows the box to fit the text. Without auto-size set
(`ppAutoSizeNone`, the default), text can overflow: shorten the text, reduce `font_size`, or grow
the shape with `shape(action: "set-size", ...)`.

