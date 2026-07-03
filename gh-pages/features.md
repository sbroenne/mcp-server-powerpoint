---
layout: default
title: "Complete Feature Reference"
description: "31 MCP tools across 10 domains for live PowerPoint automation. Slides, shapes, text, tables, charts, notes, layouts, images, and export-to-verify."
keywords: "PowerPoint MCP features, PowerPoint automation, slide tools, shape tools, chart tools, table tools, MCP operations"
permalink: /features/
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <h1 class="hero-title">Complete Feature Reference</h1>
      <p class="hero-subtitle">31 MCP tools across 10 domains for live PowerPoint automation</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">

## Tool Matrix

Every tool operates on a **session** (`session_id`) obtained from `open_presentation` or
`create_presentation`, and drives a real, live PowerPoint desktop process via COM — not an
offline `.pptx` parser.

### Presentation — session lifecycle

| Tool | What it does |
|------|---------------|
| `create_presentation` | Creates a new, blank presentation (in-memory, not yet a session) |
| `open_presentation` | Opens an existing `.pptx`/`.pptm` file and starts a session |
| `save_presentation` | Saves the active session to disk — nothing persists until this is called |
| `close_presentation` | Closes a session (optionally saving first); closing is async |
| `list_sessions` | Lists all currently open sessions, for state discovery without asking the user |

### Slide

| Tool | What it does |
|------|---------------|
| `add_slide` | Adds a new slide, optionally with a specific layout |
| `get_slide_count` | Returns the number of slides in the presentation |
| `delete_slide` | Deletes a slide by 1-based index |

### Shape

| Tool | What it does |
|------|---------------|
| `add_rectangle` | Adds a rectangle shape at a given position and size |
| `add_text_box` | Adds a text box shape |
| `get_shape_count` | Returns the number of shapes on a slide |
| `delete_shape` | Deletes a shape by 1-based index |
| `set_shape_position` | Moves a shape to a new X/Y position |
| `set_shape_size` | Resizes a shape's width/height |

### TextFrame — text content & formatting

| Tool | What it does |
|------|---------------|
| `set_text` | Sets the text content of a shape's text frame |
| `get_text` | Reads the text content of a shape's text frame |
| `set_font_size` | Sets font size for a text range |
| `set_bold` | Toggles bold for a text range |
| `set_font_color` | Sets font color for a text range |

### Table

| Tool | What it does |
|------|---------------|
| `add_table` | Adds a table shape with a given row/column count |
| `set_cell_text` | Sets the text of a table cell by 1-based row/column |
| `get_cell_text` | Reads the text of a table cell by 1-based row/column |

### Notes — speaker notes

| Tool | What it does |
|------|---------------|
| `set_notes_text` | Sets the speaker notes text for a slide |
| `get_notes_text` | Reads the speaker notes text for a slide |

### Layout

| Tool | What it does |
|------|---------------|
| `set_layout` | Applies a slide layout to a slide |
| `get_layout` | Reads the current layout of a slide |

### Image

| Tool | What it does |
|------|---------------|
| `add_picture` | Inserts a picture from a local file path onto a slide |

### Chart

| Tool | What it does |
|------|---------------|
| `add_chart` | Adds a chart shape with a given chart type and data |
| `get_chart_data` | Reads back the underlying chart data |

### Export — the visual-verification differentiator

| Tool | What it does |
|------|---------------|
| `export_slide_to_image` | Exports a single slide to an image file, for multimodal visual verification |
| `export_all_slides_to_images` | Exports every slide to image files in one call |

<div class="callout">
<strong>💡 Why export-to-verify matters</strong>
Because tools drive a real PowerPoint instance, every edit can be immediately rendered to a PNG
and checked by a vision-capable AI assistant — catching layout overlaps, overflow, and visual
regressions that text-only automation (or offline `.pptx` libraries) simply can't see.
</div>

## Design Principles

- **1-based indexing everywhere** — slide index, shape index, table row/column all start at 1,
  matching how PowerPoint itself numbers things and how a human would describe a slide deck.
- **Sessions required for edits** — `open_presentation` (or `create_presentation` + a session)
  must precede any edit tool call.
- **Explicit save** — changes exist only in the live PowerPoint process until `save_presentation`
  is called.
- **Structured results** — every tool returns `{ success, errorMessage }` (Rule 1 in the
  codebase): expected failures (bad index, missing file) come back as a clean JSON error, never
  an unhandled exception.

</div>
