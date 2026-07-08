# Canonical Workflow: Create → Build → Verify → Save → Close

The standard end-to-end loop for every PowerPoint MCP task. All 18 tools (7 session-lifecycle +
11 domain action-dispatch) exist to support this loop for one of two starting points: a brand-new
deck or an existing file.

## Starting Point A — New Presentation

```
1. create_presentation(filePath, isMacroEnabled=false) → file created on disk, NO session
2. open_presentation(filePath)                          → sessionId
3. ... build slides (see deck-builder.md) ...
4. export(action: "export-slide-to-image"/"export-all-slides-to-images", ...) → verify visually
5. save_presentation(sessionId)
6. close_presentation(sessionId)
```

## Starting Point B — Existing Presentation

```
1. open_presentation(filePath)                                     → sessionId
2. slide(action: "get-count", session_id: sessionId)                → know the current range
3. ... read/modify slides ...
4. export(action: "export-slide-to-image"/"export-all-slides-to-images", ...) → verify visually
5. save_presentation(sessionId)
6. close_presentation(sessionId)
```

## Session Management

- **One session per file, for the duration of the task.** Do not open/close the same file
  repeatedly between operations — open once, do all the work, save, close once.
- **Multiple presentations at once:** each `open_presentation` call returns an independent
  `sessionId`; pass the right one to each tool call when working across files.
- **Discover instead of asking:** `list_sessions` tells you every open session and its file path
  — use it before asking the user "which file?".
- **Always close what you open.** An unclosed session leaves a `POWERPNT.exe` process running.

## Batch Efficiency

- **Plan before executing.** For a multi-slide deck, decide the layout and content for every
  slide before calling `slide(action: "add-blank", ...)` — this avoids re-discovery mid-task (see
  `deck-builder.md`).
- **Read once, act many times.** Call `slide(action: "get-count", ...)` / `shape(action:
  "get-count", ...)` once to establish the current state, then perform the planned sequence of
  writes — don't re-query state you already know between every single write.
- **Batch text + formatting per shape.** For a given shape, call `textframe(action: "set-text",
  ...)`, then `set-font-size`/`set-bold`/`set-font-color` as needed — don't interleave unrelated
  shapes' formatting calls.
- **Save once per meaningful checkpoint**, not after every single tool call. Save after each slide
  is complete, or at the end of the whole deck for short tasks — not after every `set-text` call.

## The Discovery Actions

| Tool | Action | Use to discover |
|------|--------|------------------|
| `list_sessions` | — | Which files are currently open, and their sessionId |
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
create_presentation(filePath: "C:\Decks\q4.pptx")
open_presentation(filePath: "C:\Decks\q4.pptx") → sessionId

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
# Look at Slide1.PNG, Slide2.PNG, Slide3.PNG

save_presentation(sessionId)
close_presentation(sessionId)
```
