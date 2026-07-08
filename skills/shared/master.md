# Slide Master

Reference for `master(action: "...", ...)` — reads/edits the presentation's **slide master**:
the title and body placeholder fonts, and the master background fill color. Changes here apply
to every slide that inherits from the master (i.e. any slide that does not itself override the
property) — this is the "style the whole deck at once" tool, distinct from per-slide formatting
via `textframe`/`layout`.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `master` | `get-title-font` | `session_id` | Returns `font_name`, `font_size`, `bold`, `color_rgb` for the master's title placeholder. |
| `master` | `set-title-font` | `session_id`, `font_name?`, `font_size?`, `bold?`, `red?`, `green?`, `blue?` | Every field is optional — omit any you don't want to change. Pass `red`/`green`/`blue` together to set color. |
| `master` | `get-body-font` | `session_id` | Same shape as `get-title-font`, for the body placeholder. |
| `master` | `set-body-font` | `session_id`, `font_name?`, `font_size?`, `bold?`, `red?`, `green?`, `blue?` | Same shape as `set-title-font`, for the body placeholder. |
| `master` | `get-background-color` | `session_id` | Returns `color_rgb` for the master's background fill. |
| `master` | `set-background-color` | `session_id`, `red`, `green`, `blue` | All three color channels are required (0-255 each); sets a solid background fill. |

## What This Does — and Does Not — Cover

`master` targets the two things authors most commonly want to change across an entire deck at
once:

- The **title placeholder font** (name, size, bold, color) used by every slide's title, unless a
  slide overrides it directly.
- The **body placeholder font** (name, size, bold, color) used by every slide's body/content
  text, unless a slide overrides it directly.
- The **master background fill color** (a solid color applied behind every slide that doesn't set
  its own background).

It does **not** cover:

- Applying an entirely different theme/design — use `apply_template` (a hand-written
  session-lifecycle tool, not part of the `master` domain) to swap the whole
  masters/theme/layouts set in one call from a `.potx`/`.pptx` template file.
- Authoring or editing **custom layouts** (the individual named layouts under a master, e.g.
  "Title and Content") or adding additional slide masters — not exposed by this tool surface.

## Typical Use

Set the deck-wide look once, early, before building individual slides:

```
1. open_presentation(filePath) → sessionId
2. master(action: "set-title-font", session_id, font_name: "Segoe UI", font_size: 40, bold: true, red: 20, green: 20, blue: 20)
3. master(action: "set-body-font", session_id, font_name: "Segoe UI", font_size: 20)
4. master(action: "set-background-color", session_id, red: 255, green: 255, blue: 255)
5. ... build slides (see deck-builder.md) — inherit these fonts/background automatically ...
```

A slide only reflects the master's font/color if it hasn't overridden that property directly via
`textframe(action: "set-font-size"/"set-bold"/"set-font-color", ...)` on its own title/body shape.

## Read Before Reapplying

`master(action: "get-title-font"/"get-body-font"/"get-background-color", ...)` reports current
values — check before calling a `set-*` action if you're unsure a previous call already applied
the styling you want, rather than reapplying blindly.
