# Deck Builder: Assembling a Multi-Slide Deck

Guidance for building a complete presentation from `slide(action: "add-blank", ...)` + layout +
content tools. Activate this when the user asks to "create a deck", "build a presentation", or
"add several slides".

## Plan Before You Build

Slides only append at the end — **there is no move/reorder action in this surface.**
Decide the full slide order up front before adding slides, because re-ordering later means
deleting and re-adding slides in the right sequence.

1. `slide(action: "get-count", session_id: ...)` (or start from 0 for a new file) → know where
   you're starting from.
2. Outline every slide's purpose and layout type before creating any of them.
3. Build slides **in final order**, one at a time: `add-blank` → set layout → add content →
   verify → next slide.

## Slide Assembly Loop

Run this for every slide:

```
1. slide(action: "add-blank", session_id: ...)                                  → new blank slide appended at the end
2. layout(action: "set-layout", session_id: ..., slide_index: ..., layout_name: ...) → apply a built-in layout (see layouts.md)
3. add content: shape(action: "add-text-box"/"add-rectangle") / table(action: "add-table") / chart(action: "add-chart") / image(action: "add-picture")
4. textframe(action: "set-font-size"/"set-bold"/"set-font-color", ...) as needed (see text-formatting.md)
5. notes(action: "set-notes-text", session_id: ..., slide_index: ..., text: ...) → always add speaker notes (see speaker-notes.md)
6. export(action: "export-slide-to-image", session_id: ..., slide_index: ..., output_path: ...)  → verify (see export-and-verify.md)
7. Fix any issues found, re-verify, then move to the next slide
```

`slide(action: "add-blank", ...)` always adds a **blank** slide — all content (title, body,
shapes) is placed with `shape`/`table`/`chart`/`image` action calls afterward. There is no separate
"title+content" slide-creation action; `layout(action: "set-layout", ...)` only affects
PowerPoint's placeholder scaffolding, not what you add via shapes.

## Layout Variety

Don't build every slide the same way. Vary the composition to match content:

| Content type | Suggested composition |
|--------------|------------------------|
| Opening / section break | Large centered `shape(action: "add-text-box", ...)` title, minimal other content |
| Talking points | Title text box + one bulleted body text box (use `\n` line breaks in `text`) |
| Comparison | Title + two/three side-by-side text boxes (equal width, same top/height) |
| Metrics / trend | Title + `chart(action: "add-chart", ...)` (bar/line/pie) |
| Structured data | Title + `table(action: "add-table", ...)` |
| Product shot / diagram | Title + `image(action: "add-picture", ...)` |
| Key stat callout | Large text box with one big number, `textframe(action: "set-font-size", ..., font_size: 40)` or higher |

**Rule of thumb:** avoid using the exact same shape layout (same positions, same single text box)
for more than 2-3 consecutive slides — it reads as generated, not designed.

## Positioning Reference

`left`/`top`/`width`/`height` on `shape(action: "add-rectangle"/"add-text-box", ...)`,
`table(action: "add-table", ...)`, `chart(action: "add-chart", ...)`, and `image(action:
"add-picture", ...)` are all in **points** (PowerPoint's native unit; 1 inch = 72 points). For a
standard 16:9 slide (13.33in × 7.5in ≈ 960pt × 540pt):

- Safe margins: keep `left ≥ 40`, `top ≥ 40`, `left + width ≤ 920`, `top + height ≤ 500`.
- Title band: `top=30, height=80`.
- Body area below title: `top=130` down to `height ≈ 350`.
- Two-column split: each column `width ≈ 420` with a `40`pt gutter between them.

There is no API in this surface to query the actual slide width/height — assume the 960×540pt
16:9 default unless the user specifies otherwise, and always verify with `export(action:
"export-slide-to-image", ...)` rather than guessing blind.

## Ordering and Structure

- Title/agenda slide first, section dividers between major topics, summary/closing slide last.
- Keep related content together — build a whole section's slides consecutively rather than
  jumping around, since there's no reorder tool to fix mistakes cheaply.
- If a slide was created in the wrong position, the only fix is `slide(action: "delete", ...)` +
  re-`add-blank` in the right order — plan ahead to avoid this.

## After the Deck Is Built

1. `export(action: "export-all-slides-to-images", session_id: ..., output_directory: ...)` — one
   call renders every slide.
2. Review each exported image; fix any slide with overlapping shapes, empty placeholders, or text
   overflow (reduce `text` length or increase shape height / reduce font size).
3. `save_presentation(sessionId)`.
4. Summarize: slide count, layouts used, and the output path.
