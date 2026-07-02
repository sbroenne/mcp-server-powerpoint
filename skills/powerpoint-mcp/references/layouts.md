# Layouts

Reference for `set_layout` and `get_layout` — applies PowerPoint's built-in slide layouts by
their native `PpSlideLayout` enum member name.

## Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `set_layout` | `sessionId`, `slideIndex`, `layoutName` | `layoutName` is a `PpSlideLayout` enum member name string, e.g. `"ppLayoutBlank"`. |
| `get_layout` | `sessionId`, `slideIndex` | Returns the slide's current layout name. |

## Common Layout Names

`layoutName` must match a real `PpSlideLayout` member name exactly (case-sensitive, `pp`-prefixed
PascalCase):

| `layoutName` | Use for |
|--------------|---------|
| `ppLayoutBlank` | Fully custom slides built entirely from `add_text_box`/`add_shape`/`add_table`/`add_chart` |
| `ppLayoutTitle` | Opening/title slide |
| `ppLayoutTitleOnly` | A title band with a large open content area you'll fill yourself |
| `ppLayoutText` | Title + single text placeholder |
| `ppLayoutTwoColumnText` | Title + two text columns |
| `ppLayoutTable` | Title + table placeholder |
| `ppLayoutChart` | Title + chart placeholder |
| `ppLayoutObject` | Title + generic object placeholder (e.g., for `add_picture`) |

Passing an unrecognized string returns `success: false` — double-check spelling against the real
enum (`Microsoft.Office.Interop.PowerPoint.PpSlideLayout`) rather than guessing variants.

## What `set_layout` Actually Changes

`set_layout` changes the slide's built-in **placeholder scaffolding** (title/content
placeholders PowerPoint itself manages) — it does **not** create or move the shapes you add via
`add_text_box`/`add_table`/`add_chart`/`add_picture`. Those shapes are independent of the layout's
placeholders. In practice, for this tool surface:

- Most decks work fine with `ppLayoutBlank` on every slide and building 100% of visible content
  with the Shape/Table/Chart/Image tools directly (full control over position/size).
  `add_slide` already creates blank slides by default.
- Apply a non-blank layout (e.g. `ppLayoutTitle`) mainly when you want PowerPoint's own theme
  styling to drive title placement/formatting rather than positioning a text box manually.

## Read Before Reapplying

`get_layout` reports the current layout name — check it before calling `set_layout` if you're
unsure a previous call already applied the layout you want, rather than reapplying blindly.
