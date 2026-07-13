# Canonical Workflow: Start Session → Build → Verify → Save → Close

The standard end-to-end loop for every PowerPoint MCP task. All 13 tools and 134 operations exist
to support this loop for one of two starting points: a brand-new deck or an existing file.

## Starting Point A — New Presentation

```
1. presentation(action: "create", filePath: "C:\Decks\q4.pptx") → sessionId
2. ... build slides (see deck-builder.md) ...
3. export(action: "export-slide-to-image"/"export-all-slides-to-images", ...) → verify visually
4. presentation(action: "save", sessionId: ...)
5. presentation(action: "close", sessionId: ...)
```

## Starting Point B — Existing Presentation

```
1. presentation(action: "open", filePath: "C:\Decks\q4.pptx") → sessionId
2. slide(action: "get-count", session_id: sessionId)           → know the current range
3. ... read/modify slides ...
4. export(action: "export-slide-to-image"/"export-all-slides-to-images", ...) → verify visually
5. presentation(action: "save", sessionId: ...)
6. presentation(action: "close", sessionId: ...)
```

## Session Management

- **One session per file, for the duration of the task.** Open/create once, do all the work,
  save, close once.
- **Do not create then open the same file again.** `presentation(action: "create", ...)` already
  returns a live session.
- **Multiple presentations at once:** each `presentation(action: "open", ...)` or
  `presentation(action: "create", ...)` call returns an independent `sessionId`; pass the right one
  to each tool call when working across files.
- **Discover instead of asking:** `presentation(action: "list")` tells you every open session and
  its file path.
- **Always close what you open.** An unclosed session leaves a `POWERPNT.exe` process running.

## Batch Efficiency

- **Plan before executing.** For a multi-slide deck, decide the layout and content for every slide
  before calling `slide(action: "add-blank", ...)`.
- **Read once, act many times.** Call `slide(action: "get-count", ...)` / `shape(action:
  "get-count", ...)` once to establish the current state, then perform the planned sequence of
  writes.
- **Batch text + formatting per shape.** For a given shape, call `textframe(action: "set-text",
  ...)`, then `set-font-size`/`set-bold`/`set-font-color` as needed.
- **Save once per meaningful checkpoint**, not after every single tool call.

## The Discovery Actions

| Tool | Action | Use to discover |
|------|--------|------------------|
| `presentation` | `list` | Which files are currently open, and their `sessionId` values |
| `slide` | `get-count` | How many slides exist before adding/deleting |
| `shape` | `get-count` | How many shapes are on a slide before adding/deleting/positioning |
| `textframe` | `get-text` | Current text of a shape before editing it |
| `layout` | `get-layout` | Current layout of a slide |
| `table` | `get-cell-text` | Current content of a table cell |
| `chart` | `get-chart-data` | Category/series counts of an existing chart |
| `notes` | `get-notes-text` | Current speaker notes for a slide |

Use these instead of asking the user for information you can look up yourself (see
`behavioral-rules.md`).

## Full Example: 3-Slide Deck From Scratch

```
presentation(action: "create", filePath: "C:\Decks\q4.pptx") → sessionId

# Slide 1: title
slide(action: "add-blank", session_id: sessionId) → slideIndex=1
layout(action: "set-layout", session_id: sessionId, slide_index: 1, layout_name: "ppLayoutTitle")
shape(action: "add-text-box", session_id: sessionId, slide_index: 1, left: 50, top: 50, width: 600, height: 80, text: "Q4 Results")
textframe(action: "set-font-size", session_id: sessionId, slide_index: 1, shape_index: 1, font_size: 36)
textframe(action: "set-bold", session_id: sessionId, slide_index: 1, shape_index: 1, bold: true)
notes(action: "set-notes-text", session_id: sessionId, slide_index: 1, text: "Welcome the audience and set the scope for Q4 review.")

# Slide 2: chart
slide(action: "add-blank", session_id: sessionId) → slideIndex=2
chart(action: "add-chart", session_id: sessionId, slide_index: 2, chart_type: "bar", left: 50, top: 100, width: 500, height: 300,
      categories: ["Q1","Q2","Q3","Q4"], series_name: "Revenue", values: [120,150,170,210])
notes(action: "set-notes-text", session_id: sessionId, slide_index: 2, text: "Revenue grew steadily each quarter, accelerating in Q4.")

# Slide 3: table
slide(action: "add-blank", session_id: sessionId) → slideIndex=3
table(action: "add-table", session_id: sessionId, slide_index: 3, rows: 3, columns: 2, left: 50, top: 100, width: 400, height: 200)
table(action: "set-cell-text", session_id: sessionId, slide_index: 3, shape_index: 1, row: 1, column: 1, text: "Region")
table(action: "set-cell-text", session_id: sessionId, slide_index: 3, shape_index: 1, row: 1, column: 2, text: "Growth")
# ... remaining cells ...

# Verify
export(action: "export-all-slides-to-images", session_id: sessionId, output_directory: "C:\Decks\preview")

presentation(action: "save", sessionId: sessionId)
presentation(action: "close", sessionId: sessionId)
```
