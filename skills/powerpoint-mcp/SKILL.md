---
name: powerpoint-mcp
description: >
  PowerPoint MCP Server skill for Windows presentation automation via a live PowerPoint desktop
  instance (COM/PIA). Use when an assistant needs rich MCP tools to create, open, build, format,
  and export PowerPoint (.pptx/.pptm) presentations — slides, shapes, text boxes, tables, native
  charts, images, speaker notes, layouts, and image export for visual verification.
  Triggers: PowerPoint, presentation, deck, slides, pptx, pptm, speaker notes, chart, MCP.
compatibility: Windows with Microsoft PowerPoint desktop installed.
---

# PowerPoint MCP Server Skill

Provides 15 PowerPoint MCP tools (one presentation tool + 14 domain action-dispatch tools)
via the Model Context Protocol, driving a live PowerPoint desktop instance through the official
`Microsoft.Office.Interop.PowerPoint` PIA. Tools are auto-discovered via MCP `tools/list` — this
skill documents session lifecycle, indexing conventions, workflows, and gotchas that aren't
obvious from tool schemas alone.

Session lifecycle, Save As/copy, templates, the advisory Mark as Final flag, document properties,
and presentation tags use the `presentation` action-dispatch tool with camelCase arguments.
Domain tools (`slide`, `shape`, `textframe`, `table`, `chart`, `image`,
`notes`, `layout`, `master`, `smartart`, `animation`, `export`, `pagesetup`, `accessibility`) are
action-dispatch: one tool per domain, called as `tool(action:
"kebab-action", session_id: ..., snake_case_param: ...)`.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Create or open | `presentation(action: "create"/"open")` | Start a session, get `sessionId` | Always, before any edit |
| 3. Build | `slide(action: "add-blank")`, `shape(action: "add-rectangle"/"add-text-box"/"add-auto-shape"/"add-line"/"add-connector")`, `table(action: "add-table")`, `chart(action: "add-chart")`, `image(action: "add-picture"/"set-brightness-contrast"/"get-brightness-contrast"/"set-recolor"/"get-recolor"/"set-crop"/"get-crop")` | Add structure and content | As needed |
| 4. Format | `textframe(action: "set-font-size"/"set-bold"/"set-font-color")`, `layout(action: "set-layout")` | Apply formatting | After adding content |
| 5. Animate (optional) | `animation(action: "add-effect"/"set-transition")` | Add entrance/emphasis/exit effects or slide transitions | After content/layout are final |
| 6. Annotate | `notes(action: "set-notes-text")` | Add speaker notes | After each slide's content is final |
| 7. Verify | `export(action: "export-slide-to-image"/"export-all-slides-to-images")` | Visually confirm the result | After any visual change |
| 8. Save & close | `presentation(action: "close", save: true)` | Persist and release the session | Always last |

## Preconditions

- Windows host with Microsoft PowerPoint installed (desktop, not web/mobile).
- Use full Windows paths: `C:\Users\Name\Documents\Deck.pptx`.
- The target `.pptx`/`.pptm` file must not be open in another PowerPoint window.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: Sessions Are Required for Every Edit

Every editing action requires the `sessionId` returned by `presentation(action: "create"/"open")`.
See [Behavioral Rules](./references/behavioral-rules.md) for the full session lifecycle.

### Rule 2: Everything Is 1-Based

`slideIndex`, `shapeIndex`, and table `row`/`column` all start at 1, matching PowerPoint's own COM
object model — not 0-based like most languages. See
[Behavioral Rules](./references/behavioral-rules.md).

### Rule 3: Save-on-Close Is Explicit

Nothing is written to disk unless `presentation(action: "close", sessionId: ..., save: true)` is
used. Closing with the default `save: false` discards all changes since the last save.

### Rule 4: Close Does Not Block

`presentation(action: "close", ...)` returns immediately after removing the session; PowerPoint's own process
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
| "Which presentation is open?" | `presentation(action: "list")` |
| "How many slides are there?" | `slide(action: "get-count", session_id: sessionId)` |
| "What shapes are already on this slide?" | `shape(action: "get-count", session_id: sessionId, slide_index: slideIndex)` |

### Rule 7: Always End With a Text Summary

Never end a turn with only a tool call. State what was built, the file path, and whether it was
saved.

## Tool Selection Quick Reference

| Task | Tool(s) |
|------|---------|
| Create/open/test/close/list sessions | `presentation(action: "create"/"open"/"test"/"close"/"list")` |
| Apply template, read theme name | `presentation(action: "apply-template"/"get-theme-name")` |
| Read/set advisory Mark as Final state | `presentation(action: "get-final"/"set-final")` |
| Document metadata (built-in and custom properties) | `presentation` property actions |
| String metadata on presentations, slides, or shapes | owner-specific `set-tag`/`get-tag`/`list-tags`/`delete-tag` actions |
| Add/count/delete/duplicate/reorder slides | `slide(action: "add-blank"/"get-count"/"delete"/"duplicate"/"move-to")` |
| Per-slide background color, sections | `slide(action: "set-background-color"/"get-background-color"/"add-section"/"rename-section"/"delete-section"/"get-section-count"/"get-section-name")` |
| Add/count/delete/move/resize shapes | `shape(action: "add-rectangle"/"add-text-box"/"add-auto-shape"/"add-line"/"add-connector"/"get-count"/"delete"/"set-position"/"set-size")` |
| Format shapes and manage links | `shape(action: "set-fill"/"get-fill"/"set-line"/"get-line"/"set-rotation"/"get-rotation"/"flip"/"set-z-order"/"set-shadow"/"get-shadow"/"set-glow"/"get-glow"/"set-reflection"/"get-reflection"/"set-soft-edge"/"get-soft-edge"/"set-bevel"/"get-bevel"/"group"/"ungroup"/"set-name"/"get-name"/"set-alt-text"/"get-alt-text"/"set-hyperlink"/"get-hyperlink"/"remove-hyperlink"/"get-link-info"/"update-link"/"break-link"/"set-link-auto-update")` |
| Set/read text and font formatting | `textframe(action: "set-text"/"get-text"/"set-font-size"/"set-bold"/"set-font-color"/"set-italic"/"set-underline"/"set-font-name"/"set-alignment"/"set-bullet")` |
| Tables | `table(action: "add-table"/"set-cell-text"/"get-cell-text"/"insert-row"/"delete-row"/"insert-column"/"delete-column"/"set-cell-fill"/"get-cell-fill"/"set-cell-border"/"get-cell-border"/"merge-cells")` |
| Native charts | `chart(action: "add-chart"/"get-chart-data"/"add-series"/"replace-chart-data"/"set-chart-title"/"get-chart-title"/"set-axis-title"/"get-axis-title"/"set-legend-visibility"/"get-legend-visibility"/"set-style"/"get-style"/"set-color-style"/"get-color-style"/"set-data-table"/"get-data-table")` |
| SmartArt diagrams | `smartart(action: "add-smart-art"/"add-node"/"add-child-node"/"set-node-text"/"get-node-text"/"delete-node"/"get-node-count")` |
| Images (embedded by default; optional file links) | `image(action: "add-picture"/"set-brightness-contrast"/"get-brightness-contrast"/"set-recolor"/"get-recolor"/"set-crop"/"get-crop")` |
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
- [String tags — presentation, slide, and shape metadata](./references/tags.md)
- [Text formatting — set-text, font size/bold/color](./references/text-formatting.md)
- [Tables — add-table, cell text, row/column edits, fill/border formatting, merge](./references/tables.md)
- [Charts — add-chart/add-series categories/series/values, titles, legend](./references/charts.md)
- [SmartArt — add-smart-art layouts, node addressing, hierarchy diagrams](./references/smart-art.md)
- [Images — picture insertion and picture-format adjustments](./references/images.md)
- [Speaker notes — set/get notes](./references/speaker-notes.md)
- [Layouts — set/get slide layout](./references/layouts.md)
- [Slide master — title/body font and background color](./references/master.md)
- [Animations — entrance/emphasis/exit effects and slide transitions](./references/animations.md)
- [Export and verify — the visual verification loop](./references/export-and-verify.md)
- [Anti-patterns — common mistakes to avoid](./references/anti-patterns.md)
