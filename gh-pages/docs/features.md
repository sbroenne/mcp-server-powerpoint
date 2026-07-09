---
title: Complete Feature Reference
description: 18 MCP tools with ~98 operations across 12 domains for live PowerPoint automation. Slides, shapes, text, tables, charts, notes, layouts, masters, animations, templates, images and export-to-verify.
keywords: "PowerPoint MCP features, PowerPoint automation, slide tools, shape tools, chart tools, table tools, template tools, MCP operations"
---

# Complete Feature Reference

Every tool operates on a **session** (`session_id`) obtained from
`open_presentation` or `create_presentation`, and drives a real, live
PowerPoint desktop process via COM — not an offline `.pptx` parser.

Most domains are exposed as a single **action-dispatch tool** (e.g. `shape`,
`table`, `chart`) with an `operation` parameter selecting the specific
action — this keeps the tool list small for AI assistants while still
exposing every operation below. `Presentation` and `Template` are the
exception: they're small enough to stay as individual, hand-written tools.

## Tool matrix

### Presentation — session lifecycle

| Tool | What it does |
|------|---------------|
| `create_presentation` | Creates a new, blank presentation (in-memory, not yet a session) |
| `open_presentation` | Opens an existing `.pptx`/`.pptm` file and starts a session |
| `save_presentation` | Saves the active session to disk — nothing persists until this is called |
| `close_presentation` | Closes a session (optionally saving first); closing is async |
| `list_sessions` | Lists all currently open sessions, for state discovery without asking the user |

### Template — themes &amp; masters

| Tool | What it does |
|------|---------------|
| `apply_template` | Applies a `.potx`/`.potm`/`.pot` (or `.pptx`/`.pptm`) template's masters/theme/layouts, preserving existing slide content |
| `get_theme_name` | Reads the name of the design/theme currently applied to the presentation |

### Slide (`slide` tool — 12 operations)

| Operation | What it does |
|-----------|---------------|
| `add-blank` | Adds a new blank slide at the end of the presentation |
| `get-count` | Returns the number of slides in the presentation |
| `delete` | Deletes a slide by 1-based index |
| `duplicate` | Duplicates a slide, inserting the copy immediately after the source |
| `move-to` | Moves a slide to a new 1-based position, renumbering the rest |
| `set-background-color` | Sets a solid background color for a single slide, overriding the master |
| `get-background-color` | Reads a slide's background color and whether it follows the master |
| `add-section` | Adds a new section at a given position |
| `rename-section` | Renames a section |
| `delete-section` | Deletes a section (optionally deleting its slides too) |
| `get-section-count` | Returns the number of sections |
| `get-section-name` | Reads a section's name |

### Shape (`shape` tool — 25 operations)

| Operation | What it does |
|-----------|---------------|
| `add-rectangle` | Adds a rectangle shape at a given position and size |
| `add-text-box` | Adds a text box shape |
| `add-auto-shape` | Adds a non-rectangle auto shape (oval, arrow, star, etc.) by `MsoAutoShapeType` name |
| `add-line` | Adds a straight line between two points |
| `add-connector` | Adds a connector shape (straight, elbow, or curved) between two points |
| `get-count` | Returns the number of shapes on a slide |
| `delete` | Deletes a shape by 1-based index |
| `set-position` | Moves a shape to a new X/Y position |
| `set-size` | Resizes a shape's width/height |
| `set-fill` / `get-fill` | Sets / reads a shape's solid fill color |
| `set-line` / `get-line` | Sets / reads a shape's line/border color, weight, dash style, visibility |
| `set-rotation` / `get-rotation` | Sets / reads a shape's rotation in degrees |
| `flip` | Flips a shape horizontally or vertically |
| `set-z-order` | Moves a shape in the slide's draw order (front/back/forward/backward) |
| `set-shadow` / `get-shadow` | Toggles / reads a shape's default drop shadow |
| `group` / `ungroup` | Groups multiple shapes into one, or splits a group back apart |
| `set-name` / `get-name` | Sets / reads a shape's name (Selection Pane) |
| `set-alt-text` / `get-alt-text` | Sets / reads a shape's accessibility alt text |

### TextFrame (`textframe` tool — 15 operations)

| Operation | What it does |
|-----------|---------------|
| `set-text` / `get-text` | Sets / reads a shape's text content |
| `set-font-size` | Sets font size for a shape's text range |
| `set-bold` | Toggles bold |
| `set-font-color` | Sets font color |
| `set-italic` / `get-italic` | Sets / reads italic |
| `set-underline` / `get-underline` | Sets / reads underline |
| `set-font-name` / `get-font-name` | Sets / reads the font typeface |
| `set-alignment` / `get-alignment` | Sets / reads paragraph alignment |
| `set-bullet` / `get-bullet` | Sets / reads bullet formatting |

### Table (`table` tool — 12 operations)

| Operation | What it does |
|-----------|---------------|
| `add-table` | Adds a table shape with a given row/column count |
| `set-cell-text` / `get-cell-text` | Sets / reads a cell's text by 1-based row/column |
| `insert-row` / `delete-row` | Inserts / deletes a table row |
| `insert-column` / `delete-column` | Inserts / deletes a table column |
| `set-cell-fill` / `get-cell-fill` | Sets / reads a cell's solid fill color |
| `set-cell-border` / `get-cell-border` | Sets / reads a single cell border's color, weight, dash style, visibility |
| `merge-cells` | Merges two adjacent cells into one |

### Notes (`notes` tool — 2 operations)

| Operation | What it does |
|-----------|---------------|
| `set-notes-text` | Sets the speaker notes text for a slide |
| `get-notes-text` | Reads the speaker notes text for a slide |

### Layout (`layout` tool — 2 operations)

| Operation | What it does |
|-----------|---------------|
| `set-layout` | Applies a built-in slide layout (e.g. Title Only, Blank) |
| `get-layout` | Reads a slide's current layout name |

### Master (`master` tool — 6 operations)

| Operation | What it does |
|-----------|---------------|
| `get-title-font` / `set-title-font` | Reads / sets the slide master's title placeholder font (name, size, bold, color) |
| `get-body-font` / `set-body-font` | Reads / sets the slide master's body placeholder font |
| `get-background-color` / `set-background-color` | Reads / sets the slide master's background fill color |

Changes to the master apply to every slide that inherits from it (i.e. any slide that doesn't
override the property itself) — the practical "edit once, apply everywhere" workflow PowerPoint's
COM model supports safely.

### Animation (`animation` tool — 5 operations)

| Operation | What it does |
|-----------|---------------|
| `add-effect` | Adds an entrance/emphasis/exit animation effect to a shape by `MsoAnimEffect` name |
| `get-effect-count` | Returns the number of animation effects on a slide's timeline |
| `delete-effect` | Deletes an effect from a slide's timeline by 1-based index |
| `get-transition` / `set-transition` | Reads / sets a slide's transition (effect, duration, advance behavior) |

### Image (`image` tool — 1 operation)

| Operation | What it does |
|-----------|---------------|
| `add-picture` | Embeds a picture from a local file path onto a slide |

### Chart (`chart` tool — 9 operations)

| Operation | What it does |
|-----------|---------------|
| `add-chart` | Adds a native chart shape (bar, line, or pie) with categories and a data series |
| `get-chart-data` | Reads back a chart's category and series counts |
| `add-series` | Adds another data series to an existing chart |
| `set-chart-title` / `get-chart-title` | Sets / reads the chart's main title |
| `set-axis-title` / `get-axis-title` | Sets / reads a category or value axis title |
| `set-legend-visibility` / `get-legend-visibility` | Shows/hides / reads the chart's legend |

### Export (`export` tool — 2 operations) — the visual-verification differentiator

| Operation | What it does |
|-----------|---------------|
| `export-slide-to-image` | Exports a single slide to an image file, for multimodal visual verification |
| `export-all-slides-to-images` | Exports every slide to image files in one call |

!!! tip "Why export-to-verify matters"
    Because tools drive a real PowerPoint instance, every edit can be
    immediately rendered to a PNG and checked by a vision-capable AI
    assistant — catching layout overlaps, overflow and visual regressions
    that text-only automation (or offline `.pptx` libraries) simply can't
    see.

## Design principles

- **1-based indexing everywhere** — slide index, shape index, table
  row/column all start at 1, matching how PowerPoint itself numbers things
  and how a human would describe a slide deck.
- **Sessions required for edits** — `open_presentation` (or
  `create_presentation` + a session) must precede any edit tool call.
- **Explicit save** — changes exist only in the live PowerPoint process until
  `save_presentation` is called.
- **Structured results** — every tool returns `{ success, errorMessage }`
  (Rule 1 in the codebase): expected failures (bad index, missing file) come
  back as a clean JSON error, never an unhandled exception.
