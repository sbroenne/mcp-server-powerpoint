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

Provides 18 PowerPoint MCP tools (7 session-lifecycle tools + 11 domain action-dispatch tools)
via the Model Context Protocol, driving a live PowerPoint desktop instance through the official
`Microsoft.Office.Interop.PowerPoint` PIA. Tools are auto-discovered via MCP `tools/list` — this
skill documents session lifecycle, indexing conventions, workflows, and gotchas that aren't
obvious from tool schemas alone.

Session-lifecycle tools (`create_presentation`, `open_presentation`, `save_presentation`,
`close_presentation`, `list_sessions`, `apply_template`, `get_theme_name`) are one-tool-per-verb
with camelCase arguments. Domain tools (`slide`, `shape`, `textframe`, `table`, `chart`, `image`,
`notes`, `layout`, `master`, `animation`, `export`) are action-dispatch: one tool per domain,
called as `tool(action: "kebab-action", session_id: ..., snake_case_param: ...)`.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Create (optional) | `create_presentation` | New file on disk, no session | Only for brand-new files |
| 2. Open | `open_presentation` | Start a session, get `sessionId` | Always, before any edit |
| 3. Build | `slide(action: "add-blank")`, `shape(action: "add-rectangle"/"add-text-box"/"add-auto-shape"/"add-line"/"add-connector")`, `table(action: "add-table")`, `chart(action: "add-chart")`, `image(action: "add-picture")` | Add structure and content | As needed |
| 4. Format | `textframe(action: "set-font-size"/"set-bold"/"set-font-color")`, `layout(action: "set-layout")` | Apply formatting | After adding content |
| 5. Animate (optional) | `animation(action: "add-effect"/"set-transition")` | Add entrance/emphasis/exit effects or slide transitions | After content/layout are final |
| 6. Annotate | `notes(action: "set-notes-text")` | Add speaker notes | After each slide's content is final |
| 7. Verify | `export(action: "export-slide-to-image"/"export-all-slides-to-images")` | Visually confirm the result | After any visual change |
| 8. Save & close | `save_presentation`, `close_presentation` | Persist and release the session | Always last |

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

`export(action: "export-slide-to-image"/"export-all-slides-to-images", ...)` renders real
PowerPoint output to an image. This is the only reliable way to catch overlapping shapes, text
overflow, or chart layout problems — text-only inspection tools (`textframe(action: "get-text",
...)`, `shape(action: "get-count", ...)`) cannot. See
[Export & Verify](./references/export-and-verify.md).

### Rule 6: Never Ask Clarifying Questions

Discover state yourself instead of asking the user:

| Bad (Asking) | Good (Discovering) |
|---------------|---------------------|
| "Which presentation is open?" | `list_sessions()` |
| "How many slides are there?" | `slide(action: "get-count", session_id: sessionId)` |
| "What shapes are already on this slide?" | `shape(action: "get-count", session_id: sessionId, slide_index: slideIndex)` |

### Rule 7: Always End With a Text Summary

Never end a turn with only a tool call. State what was built, the file path, and whether it was
saved.

## Tool Selection Quick Reference

| Task | Tool(s) |
|------|---------|
| Create/open/save/close/list sessions | `create_presentation`, `open_presentation`, `save_presentation`, `close_presentation`, `list_sessions` |
| Add/count/delete slides | `slide(action: "add-blank"/"get-count"/"delete")` |
| Add/count/delete/move/resize shapes | `shape(action: "add-rectangle"/"add-text-box"/"add-auto-shape"/"add-line"/"add-connector"/"get-count"/"delete"/"set-position"/"set-size")` |
| Format shapes (fill/line/rotation/flip/z-order/shadow/group/name/alt-text) | `shape(action: "set-fill"/"get-fill"/"set-line"/"get-line"/"set-rotation"/"get-rotation"/"flip"/"set-z-order"/"set-shadow"/"get-shadow"/"group"/"ungroup"/"set-name"/"get-name"/"set-alt-text"/"get-alt-text")` |
| Set/read text and font formatting | `textframe(action: "set-text"/"get-text"/"set-font-size"/"set-bold"/"set-font-color"/"set-italic"/"set-underline"/"set-font-name"/"set-alignment"/"set-bullet")` |
| Tables | `table(action: "add-table"/"set-cell-text"/"get-cell-text")` |
| Native charts | `chart(action: "add-chart"/"get-chart-data")` |
| Images | `image(action: "add-picture")` |
| Speaker notes | `notes(action: "set-notes-text"/"get-notes-text")` |
| Slide layouts | `layout(action: "set-layout"/"get-layout")` |
| Slide master title/body font, background color | `master(action: "get-title-font"/"set-title-font"/"get-body-font"/"set-body-font"/"get-background-color"/"set-background-color")` |
| Shape entrance/emphasis/exit effects, slide transitions | `animation(action: "add-effect"/"get-effect-count"/"delete-effect"/"get-transition"/"set-transition")` |
| Visual verification | `export(action: "export-slide-to-image"/"export-all-slides-to-images")` |

## Reference Documentation

See `references/` for detailed guidance:

- [Behavioral rules — sessions, indexing, save/close semantics](./references/behavioral-rules.md)
- [Canonical create → build → verify → save → close workflow](./references/workflows.md)
- [Deck builder — assembling a multi-slide deck](./references/deck-builder.md)
- [Slides and shapes — add/position/size/delete](./references/slides-and-shapes.md)
- [Text formatting — set-text, font size/bold/color](./references/text-formatting.md)
- [Tables — add-table and cell text](./references/tables.md)
- [Charts — add-chart categories/series/values](./references/charts.md)
- [Images — add-picture](./references/images.md)
- [Speaker notes — set/get notes](./references/speaker-notes.md)
- [Layouts — set/get slide layout](./references/layouts.md)
- [Slide master — title/body font and background color](./references/master.md)
- [Animations — entrance/emphasis/exit effects and slide transitions](./references/animations.md)
- [Export and verify — the visual verification loop](./references/export-and-verify.md)
- [Anti-patterns — common mistakes to avoid](./references/anti-patterns.md)
