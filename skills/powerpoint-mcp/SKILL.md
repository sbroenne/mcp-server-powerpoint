---
name: powerpoint-mcp
description: >-
  PowerPoint MCP Server skill for Windows presentation automation via a live PowerPoint desktop
  instance (COM/PIA). Use when an assistant needs rich MCP tools to create, open, build, format,
  and export PowerPoint (.pptx/.pptm) presentations — slides, shapes, text boxes, tables, native
  charts, SmartArt, images, speaker notes, layouts, and image export for visual verification.
  Triggers: PowerPoint, presentation, deck, slides, pptx, pptm, speaker notes, chart, MCP.
---

# PowerPoint MCP Server Skill

Provides **13 PowerPoint MCP tools with 132 operations across 13 domains** via the Model Context
Protocol, driving a live PowerPoint desktop instance through the official
`Microsoft.Office.Interop.PowerPoint` PIA. Tools are auto-discovered via MCP `tools/list` — this
skill documents session lifecycle, indexing conventions, workflows, and gotchas that are not
obvious from tool schemas alone.

Every domain is a single action-dispatch tool with an `action` parameter:

- `presentation(action: "create"/"open"/"save"/"close"/"list"/...)`
- `slide(action: "add-blank"/...)`
- `shape(action: "add-rectangle"/...)`
- `textframe`, `table`, `notes`, `layout`, `master`, `animation`, `image`, `chart`, `smartart`,
  `export`

The `presentation` tool uses camelCase lifecycle/property arguments such as `filePath`, `sessionId`,
`templatePath`, `propertyName`, and `value`. The other 12 domain tools use `action` plus
`session_id` and snake_case action-specific parameters.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Start session | `presentation` | `create` for a new file, or `open` for an existing file | Always first |
| 2. Build | `slide`, `shape`, `table`, `chart`, `image`, `smartart` | Add structure and content | As needed |
| 3. Format | `textframe`, `layout`, `master`, `shape` | Apply deck, slide, and shape formatting | After adding content |
| 4. Animate (optional) | `animation` | Add effects or transitions | After content/layout are final |
| 5. Annotate | `notes` | Add speaker notes | After each slide's content is final |
| 6. Verify | `export` | Export one slide or the whole deck to images | After any visual change |
| 7. Save | `presentation` | `save` | Before finishing |
| 8. Close | `presentation` | `close` | Always last |

## Preconditions

- Windows host with Microsoft PowerPoint installed (desktop, not web/mobile).
- Use full Windows paths: `C:\Users\Name\Documents\Deck.pptx`.
- The target `.pptx`/`.pptm` file must not be open in another PowerPoint window.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: Sessions Are Required for Every Edit

Start with `presentation(action: "create", filePath: ...)` or
`presentation(action: "open", filePath: ...)`. Both return a `sessionId` used by all subsequent
edit/read operations. See [Behavioral Rules](./references/behavioral-rules.md) for the full session
lifecycle.

### Rule 2: Everything Is 1-Based

`slideIndex`, `shapeIndex`, and table `row`/`column` all start at 1, matching PowerPoint's own COM
object model — not 0-based like most languages. See
[Behavioral Rules](./references/behavioral-rules.md).

### Rule 3: Save Is Explicit

Nothing is written to disk until `presentation(action: "save", sessionId: ...)` is called.
Closing an unsaved session discards all changes since the last save.

### Rule 4: Close Does Not Block

`presentation(action: "close", sessionId: ...)` returns immediately after removing the session;
PowerPoint's own process cleanup happens afterward in the background (can take up to a few
minutes). Do not poll waiting for the OS process to exit.

### Rule 5: Verify Visually — This Is the Differentiator

`export(action: "export-slide-to-image"/"export-all-slides-to-images", ...)` renders real
PowerPoint output to an image. This is the only reliable way to catch overlapping shapes, text
overflow, or chart layout problems — text-only inspection tools cannot. See
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
| Create/open/save/close/list sessions | `presentation(action: "create"/"open"/"save"/"close"/"list", ...)` |
| Apply template, read theme name | `presentation(action: "apply-template"/"get-theme-name", ...)` |
| Document metadata (built-in and custom properties) | `presentation(action: "set-document-property"/"get-document-property"/"set-custom-property"/"get-custom-property"/"remove-custom-property", ...)` |
| Add/count/delete/duplicate/reorder slides | `slide(action: "add-blank"/"get-count"/"delete"/"duplicate"/"move-to")` |
| Per-slide background color/gradient, sections | `slide(action: "set-background-color"/"get-background-color"/"set-gradient-background"/"get-gradient-background"/"add-section"/"rename-section"/"delete-section"/"get-section-count"/"get-section-name")` |
| Add/count/delete/move/resize shapes | `shape(action: "add-rectangle"/"add-text-box"/"add-auto-shape"/"add-line"/"add-connector"/"get-count"/"delete"/"set-position"/"set-size")` |
| Format shapes and manage effects/metadata | `shape(action: "set-fill"/"get-fill"/"set-line"/"get-line"/"set-rotation"/"get-rotation"/"flip"/"set-z-order"/"set-shadow"/"get-shadow"/"set-glow"/"get-glow"/"set-reflection"/"get-reflection"/"set-soft-edge"/"get-soft-edge"/"set-bevel"/"get-bevel"/"group"/"ungroup"/"set-name"/"get-name"/"set-alt-text"/"get-alt-text"/"set-hyperlink"/"get-hyperlink"/"remove-hyperlink")` |
| Set/read text and font formatting | `textframe(action: "set-text"/"get-text"/"set-font-size"/"set-bold"/"set-font-color"/"set-italic"/"get-italic"/"set-underline"/"get-underline"/"set-font-name"/"get-font-name"/"set-alignment"/"get-alignment"/"set-bullet"/"get-bullet"/"set-auto-size"/"get-auto-size")` |
| Tables | `table(action: "add-table"/"set-cell-text"/"get-cell-text"/"insert-row"/"delete-row"/"insert-column"/"delete-column"/"set-cell-fill"/"get-cell-fill"/"set-cell-border"/"get-cell-border"/"merge-cells")` |
| Native charts | `chart(action: "add-chart"/"get-chart-data"/"add-series"/"set-chart-title"/"get-chart-title"/"set-axis-title"/"get-axis-title"/"set-legend-visibility"/"get-legend-visibility"/"replace-chart-data")` |
| SmartArt diagrams | `smartart(action: "add-smart-art"/"add-node"/"add-child-node"/"set-node-text"/"get-node-text"/"delete-node"/"get-node-count")` |
| Images | `image(action: "add-picture"/"set-brightness-contrast"/"get-brightness-contrast"/"set-recolor"/"get-recolor")` |
| Speaker notes | `notes(action: "set-notes-text"/"get-notes-text")` |
| Slide layouts | `layout(action: "set-layout"/"get-layout")` |
| Slide master title/body font and background | `master(action: "get-title-font"/"set-title-font"/"get-body-font"/"set-body-font"/"get-background-color"/"set-background-color"/"set-gradient-background"/"get-gradient-background")` |
| Shape effects and slide transitions | `animation(action: "add-effect"/"get-effect-count"/"delete-effect"/"get-transition"/"set-transition")` |
| Visual verification | `export(action: "export-slide-to-image"/"export-all-slides-to-images")` |

## Reference Documentation

See `references/` for detailed guidance:

- [Behavioral rules — sessions, indexing, save/close semantics](./references/behavioral-rules.md)
- [Canonical create/open → build → verify → save → close workflow](./references/workflows.md)
- [Deck builder — assembling a multi-slide deck](./references/deck-builder.md)
- [Slides and shapes — add/position/size/delete](./references/slides-and-shapes.md)
- [Text formatting — set-text, font size/bold/color](./references/text-formatting.md)
- [Tables — add-table, cell text, row/column edits, fill/border formatting, merge](./references/tables.md)
- [Charts — add-chart/add-series/replace-chart-data, titles, legend](./references/charts.md)
- [SmartArt — add-smart-art layouts, node addressing, hierarchy diagrams](./references/smart-art.md)
- [Images — add-picture, brightness/contrast, recolor](./references/images.md)
- [Speaker notes — set/get notes](./references/speaker-notes.md)
- [Layouts — set/get slide layout](./references/layouts.md)
- [Slide master — title/body font and background color](./references/master.md)
- [Animations — shape effects and slide transitions](./references/animations.md)
- [Export and verify — the visual verification loop](./references/export-and-verify.md)
- [Anti-patterns — common mistakes to avoid](./references/anti-patterns.md)
