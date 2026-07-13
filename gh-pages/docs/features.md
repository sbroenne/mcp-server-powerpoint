---
title: Complete Feature Reference
description: 13 MCP tools with 134 operations across 13 domains for live PowerPoint automation through single action-dispatch tools.
keywords: "PowerPoint MCP features, PowerPoint automation, presentation tool, slide tool, shape tool, chart tool, SmartArt tool, export-to-verify"
---

# Complete Feature Reference

PowerPoint MCP Server exposes **13 MCP tools with 134 operations across 13 domains**.
Every domain is a **single action-dispatch tool** that takes an `action` parameter — for example
`presentation(action="open", filePath="C:\\Decks\\q4.pptx")` or
`chart(action="add-chart", session_id="...", slide_index=2, ...)`.

The CLI mirrors the same domain model:

- `pptcli session <action> ...` for the `presentation` domain's session/template/property work
- `pptcli <domain> <action> ...` for all other domains, such as `pptcli chart add-chart ...`

## Tool matrix

| Tool | Ops | What it covers | MCP call shape | CLI shape |
|------|-----|----------------|----------------|-----------|
| `presentation` | 12 | Session lifecycle, template application, built-in/custom document properties | `presentation(action="...", ...)` | `pptcli session <action> ...` |
| `slide` | 14 | Slide lifecycle, slide backgrounds, sections | `slide(action="...", session_id=..., ...)` | `pptcli slide <action> -s <SESSION_ID> ...` |
| `shape` | 36 | Shapes, geometry, styling, effects, grouping, naming, hyperlinks | `shape(action="...", session_id=..., ...)` | `pptcli shape <action> -s <SESSION_ID> ...` |
| `textframe` | 17 | Text content and text formatting | `textframe(action="...", session_id=..., ...)` | `pptcli textframe <action> -s <SESSION_ID> ...` |
| `table` | 12 | Table creation and cell editing/formatting | `table(action="...", session_id=..., ...)` | `pptcli table <action> -s <SESSION_ID> ...` |
| `notes` | 2 | Speaker notes | `notes(action="...", session_id=..., ...)` | `pptcli notes <action> -s <SESSION_ID> ...` |
| `layout` | 2 | Slide layouts | `layout(action="...", session_id=..., ...)` | `pptcli layout <action> -s <SESSION_ID> ...` |
| `master` | 8 | Slide master fonts and backgrounds | `master(action="...", session_id=..., ...)` | `pptcli master <action> -s <SESSION_ID> ...` |
| `animation` | 5 | Shape effects and slide transitions | `animation(action="...", session_id=..., ...)` | `pptcli animation <action> -s <SESSION_ID> ...` |
| `image` | 7 | Picture insertion and picture adjustments (brightness/contrast, recolor, crop) | `image(action="...", session_id=..., ...)` | `pptcli image <action> -s <SESSION_ID> ...` |
| `chart` | 10 | Native charts, titles, axes, legend, data replacement | `chart(action="...", session_id=..., ...)` | `pptcli chart <action> -s <SESSION_ID> ...` |
| `smartart` | 7 | SmartArt diagrams and node editing | `smartart(action="...", session_id=..., ...)` | `pptcli smartart <action> -s <SESSION_ID> ...` |
| `export` | 2 | Export-to-verify image rendering | `export(action="...", session_id=..., ...)` | `pptcli export <action> -s <SESSION_ID> ...` |

## Domain reference

### `presentation` tool (12 operations)

Use `presentation` for session lifecycle, templates/themes, and document properties.
`create` and `open` establish a session and return a `sessionId`; the remaining edit/read actions
use that `sessionId`.

| Action | What it does |
|--------|---------------|
| `create` | Create a new presentation file, save it immediately, and leave the session open. |
| `open` | Open an existing presentation file and start a session. |
| `save` | Save the presentation for an open session. |
| `close` | Close an open session; PowerPoint shutdown continues in the background. |
| `list` | List all currently open sessions. |
| `apply-template` | Apply a `.potx`/`.potm`/`.pot` or `.pptx`/`.pptm` template source while preserving slide content. |
| `get-theme-name` | Read the currently applied design/theme name. |
| `set-document-property` | Set a built-in document metadata property such as Title or Author. |
| `get-document-property` | Read a built-in document metadata property. |
| `set-custom-property` | Create or update a custom document property. |
| `get-custom-property` | Read a custom document property. |
| `remove-custom-property` | Remove a custom document property. |

**Exact action order:** `create`, `open`, `save`, `close`, `list`, `apply-template`,
`get-theme-name`, `set-document-property`, `get-document-property`, `set-custom-property`,
`get-custom-property`, `remove-custom-property`

### `slide` tool (14 operations)

| Action | What it does |
|--------|---------------|
| `add-blank` | Add a blank slide. |
| `get-count` | Return the slide count. |
| `delete` | Delete a slide by 1-based index. |
| `duplicate` | Duplicate a slide. |
| `move-to` | Move a slide to a new 1-based position. |
| `set-background-color` | Set a slide's solid background color. |
| `get-background-color` | Read a slide's solid background color / master-follow state. |
| `set-gradient-background` | Set a slide's gradient background. |
| `get-gradient-background` | Read a slide's gradient background. |
| `add-section` | Add a section. |
| `rename-section` | Rename a section. |
| `delete-section` | Delete a section. |
| `get-section-count` | Return the section count. |
| `get-section-name` | Read a section name. |

**Exact action order:** `add-blank`, `get-count`, `delete`, `duplicate`, `move-to`,
`set-background-color`, `get-background-color`, `set-gradient-background`,
`get-gradient-background`, `add-section`, `rename-section`, `delete-section`,
`get-section-count`, `get-section-name`

### `shape` tool (36 operations)

Use `shape` for shape creation, geometry, styling, effects, grouping, naming, alt text, and
hyperlinks.

**Exact action order:** `add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`,
`add-connector`, `get-count`, `delete`, `set-position`, `set-size`, `set-fill`, `get-fill`,
`set-line`, `get-line`, `set-rotation`, `get-rotation`, `flip`, `set-z-order`, `set-shadow`,
`get-shadow`, `set-glow`, `get-glow`, `set-reflection`, `get-reflection`, `set-soft-edge`,
`get-soft-edge`, `set-bevel`, `get-bevel`, `group`, `ungroup`, `set-name`, `get-name`,
`set-alt-text`, `get-alt-text`, `set-hyperlink`, `get-hyperlink`, `remove-hyperlink`

### `textframe` tool (17 operations)

Use `textframe` for text content and font/paragraph formatting on a shape's text frame.

**Exact action order:** `set-text`, `get-text`, `set-font-size`, `set-bold`, `set-font-color`,
`set-italic`, `get-italic`, `set-underline`, `get-underline`, `set-font-name`, `get-font-name`,
`set-alignment`, `get-alignment`, `set-bullet`, `get-bullet`, `set-auto-size`, `get-auto-size`

### `table` tool (12 operations)

Use `table` for native PowerPoint tables.

**Exact action order:** `add-table`, `set-cell-text`, `get-cell-text`, `insert-row`, `delete-row`,
`insert-column`, `delete-column`, `set-cell-fill`, `get-cell-fill`, `set-cell-border`,
`get-cell-border`, `merge-cells`

### `notes` tool (2 operations)

**Exact action order:** `set-notes-text`, `get-notes-text`

### `layout` tool (2 operations)

**Exact action order:** `set-layout`, `get-layout`

### `master` tool (8 operations)

Use `master` for deck-wide master placeholder fonts and master backgrounds.

**Exact action order:** `get-title-font`, `set-title-font`, `get-body-font`, `set-body-font`,
`get-background-color`, `set-background-color`, `set-gradient-background`,
`get-gradient-background`

### `animation` tool (5 operations)

**Exact action order:** `add-effect`, `get-effect-count`, `delete-effect`, `get-transition`,
`set-transition`

### `image` tool (7 operations)

Use `image` for inserting and adjusting pictures.

**Exact action order:** `add-picture`, `set-brightness-contrast`, `get-brightness-contrast`,
`set-recolor`, `get-recolor`, `set-crop`, `get-crop`

### `chart` tool (10 operations)

Use `chart` for native PowerPoint charts.

**Exact action order:** `add-chart`, `get-chart-data`, `add-series`, `set-chart-title`,
`get-chart-title`, `set-axis-title`, `get-axis-title`, `set-legend-visibility`,
`get-legend-visibility`, `replace-chart-data`

### `smartart` tool (7 operations)

Use `smartart` for SmartArt diagrams and node editing.

**Exact action order:** `add-smart-art`, `add-node`, `add-child-node`, `set-node-text`,
`get-node-text`, `delete-node`, `get-node-count`

### `export` tool (2 operations)

Use `export` for the project's export-to-verify loop.

**Exact action order:** `export-slide-to-image`, `export-all-slides-to-images`

!!! tip "Why export-to-verify matters"
    Because the tools drive a **real PowerPoint desktop instance**, every visual edit can be
    rendered to an image and checked by a vision-capable AI assistant before the deck is declared
    done.

## Design principles

- **Single action-dispatch tool per domain** — fewer MCP tools, clearer schemas, same total power.
- **`action`, not `operation`** — every MCP domain tool selects its behavior with an `action`
  enum parameter.
- **1-based indexing everywhere** — slides, shapes, rows, and columns all match PowerPoint's own
  object model.
- **Sessions are explicit** — open/create once, do the work, save, then close.
- **Export to verify** — when a change is visual, render it and inspect the image.
