# Layouts

Reference for slide layouts and presentation-wide page setup. `layout` applies PowerPoint's native
slide layouts; `pagesetup` controls slide dimensions, numbering, and footer elements.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `layout` | `set-layout` | `session_id`, `slide_index`, `layout_name` | `layout_name` is a `PpSlideLayout` enum member name string, e.g. `"ppLayoutBlank"`. |
| `layout` | `get-layout` | `session_id`, `slide_index` | Returns the slide's current layout name. |
| `layout` | `list-layouts` | `session_id`, `master_index` | Lists custom layouts for a 1-based slide master. |
| `layout` | `delete-layout` | `session_id`, `master_index`, `layout_index` | Deletes an unused custom layout; PowerPoint refuses layouts still used by slides. |
| `pagesetup` | `get-settings` | `session_id` | Returns slide width/height in points, orientation, and first slide number. |
| `pagesetup` | `set-size` | `session_id`, `width`, `height` | Sets custom slide dimensions in points. |
| `pagesetup` | `set-first-slide-number` | `session_id`, `first_slide_number` | Sets the number displayed on the first slide. |
| `pagesetup` | `get-footer` | `session_id` | Reads footer text, slide-number visibility, and date/time settings. |
| `pagesetup` | `set-footer` | `session_id` plus optional footer fields | Changes only the supplied footer fields, including `show_footer`, slide numbers, date/time, and title-slide display. `date_time_mode` is `automatic` or `fixed`. |

## Common Layout Names

`layout_name` must match a real `PpSlideLayout` member name exactly (case-sensitive, `pp`-prefixed
PascalCase):

| `layout_name` | Use for |
|--------------|---------|
| `ppLayoutBlank` | Fully custom slides built entirely from `shape`/`table`/`chart` action calls |
| `ppLayoutTitle` | Opening/title slide |
| `ppLayoutTitleOnly` | A title band with a large open content area you'll fill yourself |
| `ppLayoutText` | Title + single text placeholder |
| `ppLayoutTwoColumnText` | Title + two text columns |
| `ppLayoutTable` | Title + table placeholder |
| `ppLayoutChart` | Title + chart placeholder |
| `ppLayoutObject` | Title + generic object placeholder (e.g., for `image(action: "add-picture", ...)`) |

Passing an unrecognized string returns `success: false` — double-check spelling against the real
enum (`Microsoft.Office.Interop.PowerPoint.PpSlideLayout`) rather than guessing variants.

## What `set-layout` Actually Changes

`layout(action: "set-layout", ...)` changes the slide's built-in **placeholder scaffolding**
(title/content placeholders PowerPoint itself manages) — it does **not** create or move the shapes
you add via `shape`/`table`/`chart`/`image` action calls. Those shapes are independent of the
layout's placeholders. In practice, for this tool surface:

- Most decks work fine with `ppLayoutBlank` on every slide and building 100% of visible content
  with the Shape/Table/Chart/Image tools directly (full control over position/size).
  `slide(action: "add-blank", ...)` already creates blank slides by default.
- Apply a non-blank layout (e.g. `ppLayoutTitle`) mainly when you want PowerPoint's own theme
  styling to drive title placement/formatting rather than positioning a text box manually.

## Read Before Reapplying

`layout(action: "get-layout", ...)` reports the current layout name — check it before calling
`set-layout` if you're unsure a previous call already applied the layout you want, rather than
reapplying blindly.
