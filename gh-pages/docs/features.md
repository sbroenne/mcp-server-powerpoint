---
title: Complete Feature Reference
description: 15 MCP tools with 160 operations across 15 domains for live PowerPoint automation through single action-dispatch tools.
keywords: "PowerPoint MCP features, PowerPoint automation, presentation tool, slide tool, shape tool, chart tool, SmartArt tool, export-to-verify"
---

# Complete Feature Reference

PowerPoint MCP Server exposes **15 MCP tools with 160 operations across 15 domains**.
Every domain is a **single action-dispatch tool** that takes an `action` parameter — for example
`presentation(action="open", filePath="C:\\Decks\\q4.pptx")` or
`chart(action="add-chart", session_id="...", slide_index=2, ...)`.

The CLI mirrors the same domain model:

- `pptcli session <action> ...` for the `presentation` domain's session/template/property work
- `pptcli <domain> <action> ...` for all other domains, such as `pptcli chart add-chart ...`

## Tool matrix

| Tool | Ops | What it covers | MCP call shape | CLI shape |
|------|-----|----------------|----------------|-----------|
| `presentation` | 14 | Session lifecycle, Save As/copy, open testing, template application, built-in/custom document properties | `presentation(action="...", ...)` | `pptcli session <action> ...` |
| `slide` | 19 | Slide lifecycle, backgrounds, sections, comments, import | `slide(action="...", session_id=..., ...)` | `pptcli slide <action> -s <SESSION_ID> ...` |
| `shape` | 39 | Shapes, styling, grouping, hyperlinks, placeholders | `shape(action="...", session_id=..., ...)` | `pptcli shape <action> -s <SESSION_ID> ...` |
| `textframe` | 20 | Text content and text formatting | `textframe(action="...", session_id=..., ...)` | `pptcli textframe <action> -s <SESSION_ID> ...` |
| `table` | 12 | Table creation and cell editing/formatting | `table(action="...", session_id=..., ...)` | `pptcli table <action> -s <SESSION_ID> ...` |
| `notes` | 2 | Speaker notes | `notes(action="...", session_id=..., ...)` | `pptcli notes <action> -s <SESSION_ID> ...` |
| `layout` | 4 | Slide layouts | `layout(action="...", session_id=..., ...)` | `pptcli layout <action> -s <SESSION_ID> ...` |
| `pagesetup` | 5 | Slide size, numbering, footer, date/time | `pagesetup(action="...", session_id=..., ...)` | `pptcli pagesetup <action> -s <SESSION_ID> ...` |
| `accessibility` | 3 | Deterministic audit and reading order | `accessibility(action="...", session_id=..., ...)` | `pptcli accessibility <action> -s <SESSION_ID> ...` |
| `master` | 10 | Slide master fonts and backgrounds | `master(action="...", session_id=..., ...)` | `pptcli master <action> -s <SESSION_ID> ...` |
| `animation` | 5 | Shape effects and slide transitions | `animation(action="...", session_id=..., ...)` | `pptcli animation <action> -s <SESSION_ID> ...` |
| `image` | 7 | Picture insertion and picture adjustments (brightness/contrast, recolor, crop) | `image(action="...", session_id=..., ...)` | `pptcli image <action> -s <SESSION_ID> ...` |
| `chart` | 10 | Native charts, titles, axes, legend, data replacement | `chart(action="...", session_id=..., ...)` | `pptcli chart <action> -s <SESSION_ID> ...` |
| `smartart` | 7 | SmartArt diagrams and node editing | `smartart(action="...", session_id=..., ...)` | `pptcli smartart <action> -s <SESSION_ID> ...` |
| `export` | 3 | PDF delivery and export-to-verify image rendering | `export(action="...", session_id=..., ...)` | `pptcli export <action> -s <SESSION_ID> ...` |

## Domain reference

### `presentation` tool (14 operations)

Use `presentation` for session lifecycle, Save As/copy, templates/themes, and document properties.
`create` and `open` establish a session and return a `sessionId`; the remaining edit/read actions
use that `sessionId`.

| Action | What it does |
|--------|---------------|
| `create` | Create a new presentation file, save it immediately, and leave the session open. |
| `open` | Open an existing presentation file and start a session. |
| `close` | Close an open session, optionally saving first; PowerPoint shutdown continues in the background. |
| `list` | List all currently open sessions. |
| `test` | Check whether PowerPoint can open a file without retaining a live session. |
| `save-as` | Save the active presentation under a new `.pptx`, `.pptm`, or `.ppt` path and update the session to that path. Existing files require `overwrite=true`. |
| `save-copy-as` | Save a copy using the active presentation's current format without changing the active presentation or session path. Existing files require `overwrite=true`. |
| `apply-template` | Apply a `.potx`/`.potm`/`.pot` or `.pptx`/`.pptm` template source while preserving slide content. |
| `get-theme-name` | Read the currently applied design/theme name. |
| `set-document-property` | Set a built-in document metadata property such as Title or Author. |
| `get-document-property` | Read a built-in document metadata property. |
| `set-custom-property` | Create or update a custom document property. |
| `get-custom-property` | Read a custom document property. |
| `remove-custom-property` | Remove a custom document property. |

**Exact action order:** `create`, `open`, `close`, `list`, `test`, `save-as`, `save-copy-as`,
`apply-template`, `get-theme-name`, `set-document-property`, `get-document-property`,
`set-custom-property`, `get-custom-property`, `remove-custom-property`

### `slide` tool (19 operations)

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
| `list-comments` | List legacy slide comments exposed by the native PowerPoint COM API. |
| `add-comment` | Add a legacy comment to a slide. |
| `delete-comment` | Delete a legacy comment by 1-based index. |
| `clear-comments` | Remove all legacy comments from a slide. |
| `import-from-file` | Insert a 1-based source slide range after a destination slide. |

**Exact action order:** `add-blank`, `get-count`, `delete`, `duplicate`, `move-to`,
`set-background-color`, `get-background-color`, `set-gradient-background`,
`get-gradient-background`, `add-section`, `rename-section`, `delete-section`,
`get-section-count`, `get-section-name`, `list-comments`, `add-comment`, `delete-comment`,
`clear-comments`, `import-from-file`

### `shape` tool (39 operations)

Use `shape` for shape creation, geometry, styling, effects, grouping, naming, alt text, and
hyperlinks.

**Exact action order:** `add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`,
`add-connector`, `get-count`, `delete`, `set-position`, `set-size`, `set-fill`, `get-fill`,
`set-line`, `get-line`, `set-rotation`, `get-rotation`, `flip`, `set-z-order`, `set-shadow`,
`get-shadow`, `set-glow`, `get-glow`, `set-reflection`, `get-reflection`, `set-soft-edge`,
`get-soft-edge`, `set-bevel`, `get-bevel`, `group`, `ungroup`, `set-name`, `get-name`,
`set-alt-text`, `get-alt-text`, `set-hyperlink`, `get-hyperlink`, `remove-hyperlink`,
`list-placeholders`, `set-placeholder-text`, `set-placeholder-image`

### `textframe` tool (20 operations)

Use `textframe` for text content and font/paragraph formatting on a shape's text frame.

**Exact action order:** `set-text`, `get-text`, `set-font-size`, `get-font-size`, `set-bold`,
`get-bold`, `set-font-color`, `get-font-color`, `set-italic`, `get-italic`, `set-underline`,
`get-underline`, `set-font-name`, `get-font-name`, `set-alignment`, `get-alignment`, `set-bullet`,
`get-bullet`, `set-auto-size`, `get-auto-size`

### `table` tool (12 operations)

Use `table` for native PowerPoint tables.

**Exact action order:** `add-table`, `set-cell-text`, `get-cell-text`, `insert-row`, `delete-row`,
`insert-column`, `delete-column`, `set-cell-fill`, `get-cell-fill`, `set-cell-border`,
`get-cell-border`, `merge-cells`

### `notes` tool (2 operations)

**Exact action order:** `set-notes-text`, `get-notes-text`

### `layout` tool (4 operations)

**Exact action order:** `set-layout`, `get-layout`, `list-layouts`, `delete-layout`

### `pagesetup` tool (5 operations)

Use `pagesetup` for presentation-wide slide dimensions, numbering, and footer settings.

**Exact action order:** `get-settings`, `set-size`, `set-first-slide-number`, `get-footer`,
`set-footer`

### `accessibility` tool (3 operations)

Use `accessibility` for deterministic checks and slide reading order. The audit covers missing
alternative text on visual content and empty title placeholders; it is not an AI writing review.

**Exact action order:** `audit`, `get-reading-order`, `set-reading-order`

### `master` tool (10 operations)

Use `master` for deck-wide master placeholder fonts and master backgrounds.

**Exact action order:** `get-title-font`, `set-title-font`, `get-body-font`, `set-body-font`,
`get-background-color`, `set-background-color`, `set-gradient-background`,
`get-gradient-background`, `list-masters`, `delete-master`

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

### `export` tool (3 operations)

Use `export` for the project's export-to-verify loop.

**Exact action order:** `export-to-pdf`, `export-slide-to-image`,
`export-all-slides-to-images`

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
- **Sessions are explicit** — open/create once, do the work, then close with the chosen save behavior.
- **Export to verify** — when a change is visual, render it and inspect the image.
