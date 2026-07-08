# Export & Visual Verification

Reference for `export(action: "export-slide-to-image", ...)` and `export(action:
"export-all-slides-to-images", ...)` — PowerPoint's native COM rendering used as the multimodal
"look at the result" verification loop. This is the tool surface's differentiator: text-only
inspection (`textframe(action: "get-text", ...)`, `shape(action: "get-count", ...)`,
`chart(action: "get-chart-data", ...)`) cannot catch overlapping shapes, text overflow, bad chart
proportions, or wrong colors — only a rendered image can.

## REQUIRED: Verify After Visual Changes

**You MUST export and look at the result after creating or repositioning any visual content** —
shapes, tables, charts, images, or significant text/formatting changes. Do not save and close a
session without this step when the task involves visual output.

```
1. chart(action: "add-chart", ...) / table(action: "add-table", ...) / image(action: "add-picture", ...) / shape(action: "set-position", ...)
2. export(action: "export-slide-to-image", session_id: ..., slide_index: ..., output_path: ...)  ← REQUIRED — never skip
3. Inspect the returned image for overlap, overflow, or wrong placement
4. If issues found → fix → export again → repeat until it looks right
5. save_presentation(sessionId)
```

This rule applies even if the operation reported `success: true` — a successful COM call only
confirms the API accepted the parameters, not that the result looks correct.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `export` | `export-slide-to-image` | `session_id`, `slide_index`, `output_path`, `format` (default `"PNG"`), `width`, `height` (optional pixels) | Renders exactly one slide to a single image file. |
| `export` | `export-all-slides-to-images` | `session_id`, `output_directory`, `format` (default `"PNG"`) | Renders every slide; PowerPoint names files `Slide1.PNG`, `Slide2.PNG`, etc. in the given directory. |

- `format` accepts any PowerPoint export filter name: `"PNG"`, `"JPG"`, `"GIF"`, `"BMP"`, `"TIF"`,
  `"WMF"`, `"EMF"`. Default to `"PNG"` unless the user needs a specific format.
- `width`/`height` on `export-slide-to-image` control output pixel dimensions; omit them to use
  PowerPoint's default rendering size.
- `export-all-slides-to-images` creates `output_directory` if it doesn't already exist.

## When to Use Each

| Situation | Use |
|-----------|-----|
| Just added/changed one slide | `export(action: "export-slide-to-image", ...)` on that slide only |
| Finished building a whole deck | `export(action: "export-all-slides-to-images", ...)` once, review every image |
| Iterating on a single slide's layout | `export-slide-to-image` repeatedly on that slide during the fix loop |

Prefer the single-slide export while iterating on one slide — exporting the whole deck on every
fix cycle wastes calls once you've localized the issue to one slide.

## Common Issues to Look For

| Problem | Likely fix |
|---------|-----------|
| Text visually cut off / overflowing its box | Shorten `text`, reduce `font_size`, or grow the shape with `shape(action: "set-size", ...)` |
| Two shapes overlapping | `shape(action: "set-position", ...)` on one of them, using the positioning reference in `deck-builder.md` |
| Chart illegible / too small | Increase `width`/`height` on the chart, or reduce category count |
| Table cell text overflowing | Shorten cell text or increase the table's `height` when creating it (tables can't be resized post-creation without re-adding — see `tables.md`) |
| Image stretched/squashed | Recompute `width`/`height` to match the source image's aspect ratio |

## The Fix Loop

```
For each slide with visual content {
  1. Build/modify the slide
  2. export(action: "export-slide-to-image", ...) → inspect
  3. If ANY issue found → fix it → export again
  4. Move to the next slide only when it looks right
}
```

Expect 1-2 fix cycles per visually complex slide (charts, tables, multi-shape layouts) — this is
normal, not a sign something went wrong the first time.

## After the Full Deck

Before the final `save_presentation` + `close_presentation`, run `export(action:
"export-all-slides-to-images", ...)` once as a final pass over the whole deck, confirming no slide
was missed and the deck reads coherently start to finish.
