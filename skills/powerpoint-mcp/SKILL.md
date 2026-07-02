---
name: powerpoint-mcp
description: >
  PowerPoint MCP Server skill for Windows presentation automation via a live PowerPoint desktop
  instance (COM/PIA). Use when an assistant needs rich MCP tools to create, open, build, format,
  and export PowerPoint (.pptx/.pptm) presentations — slides, shapes, text boxes, tables, native
  charts, images, speaker notes, layouts, and image export for visual verification.
  Triggers: PowerPoint, presentation, deck, slides, pptx, pptm, speaker notes, chart, MCP.
---

# PowerPoint MCP Server Skill

Provides 31 PowerPoint operations via the Model Context Protocol, driving a live PowerPoint
desktop instance through the official `Microsoft.Office.Interop.PowerPoint` PIA. Tools are
auto-discovered via MCP `tools/list` — this skill documents session lifecycle, indexing
conventions, workflows, and gotchas that aren't obvious from tool schemas alone.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Create (optional) | `create_presentation` | New file on disk, no session | Only for brand-new files |
| 2. Open | `open_presentation` | Start a session, get `sessionId` | Always, before any edit |
| 3. Build | `add_slide`, `add_rectangle`/`add_text_box`/`add_table`/`add_chart`/`add_picture` | Add structure and content | As needed |
| 4. Format | `set_font_size`/`set_bold`/`set_font_color`, `set_layout` | Apply formatting | After adding content |
| 5. Annotate | `set_notes_text` | Add speaker notes | After each slide's content is final |
| 6. Verify | `export_slide_to_image` / `export_all_slides_to_images` | Visually confirm the result | After any visual change |
| 7. Save & close | `save_presentation`, `close_presentation` | Persist and release the session | Always last |

## Preconditions

- Windows host with Microsoft PowerPoint installed (desktop, not web/mobile).
- Use full Windows paths: `C:\Users\Name\Documents\Deck.pptx`.
- The target `.pptx`/`.pptm` file must not be open in another PowerPoint window.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: Sessions Are Required for Every Edit

All tools except `create_presentation` require a `sessionId` from `open_presentation`. See
[Behavioral Rules](./references/behavioral-rules.md) for the full session lifecycle, including
why `create_presentation` deliberately does NOT leave a session open.

### Rule 2: Everything Is 1-Based

`slideIndex`, `shapeIndex`, and table `row`/`column` all start at 1, matching PowerPoint's own COM
object model — not 0-based like most languages. See
[Behavioral Rules](./references/behavioral-rules.md).

### Rule 3: Save Is Explicit

Nothing is written to disk until `save_presentation(sessionId)` is called. Closing an unsaved
session discards all changes since the last save.

### Rule 4: Close Does Not Block

`close_presentation` returns immediately after removing the session; PowerPoint's own process
cleanup happens afterward in the background (can take up to a few minutes). Do not poll waiting
for the OS process to exit.

### Rule 5: Verify Visually — This Is the Differentiator

`export_slide_to_image` / `export_all_slides_to_images` render real PowerPoint output to an
image. This is the only reliable way to catch overlapping shapes, text overflow, or chart layout
problems — text-only inspection tools (`get_text`, `get_shape_count`) cannot. See
[Export & Verify](./references/export-and-verify.md).

### Rule 6: Never Ask Clarifying Questions

Discover state yourself instead of asking the user:

| Bad (Asking) | Good (Discovering) |
|---------------|---------------------|
| "Which presentation is open?" | `list_sessions()` |
| "How many slides are there?" | `get_slide_count(sessionId)` |
| "What shapes are already on this slide?" | `get_shape_count(sessionId, slideIndex)` |

### Rule 7: Always End With a Text Summary

Never end a turn with only a tool call. State what was built, the file path, and whether it was
saved.

## Tool Selection Quick Reference

| Task | Tool(s) |
|------|---------|
| Create/open/save/close/list sessions | `create_presentation`, `open_presentation`, `save_presentation`, `close_presentation`, `list_sessions` |
| Add/count/delete slides | `add_slide`, `get_slide_count`, `delete_slide` |
| Add/count/delete/move/resize shapes | `add_rectangle`, `add_text_box`, `get_shape_count`, `delete_shape`, `set_shape_position`, `set_shape_size` |
| Set/read text and font formatting | `set_text`, `get_text`, `set_font_size`, `set_bold`, `set_font_color` |
| Tables | `add_table`, `set_cell_text`, `get_cell_text` |
| Native charts | `add_chart`, `get_chart_data` |
| Images | `add_picture` |
| Speaker notes | `set_notes_text`, `get_notes_text` |
| Slide layouts | `set_layout`, `get_layout` |
| Visual verification | `export_slide_to_image`, `export_all_slides_to_images` |

## Reference Documentation

See `references/` for detailed guidance:

- [Behavioral rules — sessions, indexing, save/close semantics](./references/behavioral-rules.md)
- [Canonical create → build → verify → save → close workflow](./references/workflows.md)
- [Deck builder — assembling a multi-slide deck](./references/deck-builder.md)
- [Slides and shapes — add/position/size/delete](./references/slides-and-shapes.md)
- [Text formatting — set_text, font size/bold/color](./references/text-formatting.md)
- [Tables — add_table and cell text](./references/tables.md)
- [Charts — add_chart categories/series/values](./references/charts.md)
- [Images — add_picture](./references/images.md)
- [Speaker notes — set/get notes](./references/speaker-notes.md)
- [Layouts — set/get slide layout](./references/layouts.md)
- [Export and verify — the visual verification loop](./references/export-and-verify.md)
- [Anti-patterns — common mistakes to avoid](./references/anti-patterns.md)
