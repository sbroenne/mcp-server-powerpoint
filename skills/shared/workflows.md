# Canonical Workflow: Create → Build → Verify → Save → Close

The standard end-to-end loop for every PowerPoint MCP task. All 31 tools exist to support this
loop for one of two starting points: a brand-new deck or an existing file.

## Starting Point A — New Presentation

```
1. create_presentation(filePath, isMacroEnabled=false) → file created on disk, NO session
2. open_presentation(filePath)                          → sessionId
3. ... build slides (see deck-builder.md) ...
4. export_slide_to_image / export_all_slides_to_images   → verify visually
5. save_presentation(sessionId)
6. close_presentation(sessionId)
```

## Starting Point B — Existing Presentation

```
1. open_presentation(filePath)                          → sessionId
2. get_slide_count(sessionId)                            → know the current range
3. ... read/modify slides ...
4. export_slide_to_image / export_all_slides_to_images   → verify visually
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
  slide before calling `add_slide` — this avoids re-discovery mid-task (see
  `deck-builder.md`).
- **Read once, act many times.** Call `get_slide_count` / `get_shape_count` once to establish the
  current state, then perform the planned sequence of writes — don't re-query state you already
  know between every single write.
- **Batch text + formatting per shape.** For a given shape, call `set_text`, then
  `set_font_size`/`set_bold`/`set_font_color` as needed — don't interleave unrelated shapes'
  formatting calls.
- **Save once per meaningful checkpoint**, not after every single tool call. Save after each slide
  is complete, or at the end of the whole deck for short tasks — not after every `set_text` call.

## The Discovery Tools

| Tool | Use to discover |
|------|------------------|
| `list_sessions` | Which files are currently open, and their sessionId |
| `get_slide_count` | How many slides exist before adding/deleting |
| `get_shape_count` | How many shapes are on a slide before adding/deleting/positioning |
| `get_text` | Current text of a shape before editing it |
| `get_layout` | Current layout of a slide |
| `get_cell_text` | Current content of a table cell |
| `get_chart_data` | Category/series counts of an existing chart |
| `get_notes_text` | Current speaker notes for a slide |

Use these instead of asking the user for information you can look up yourself (see
`behavioral-rules.md`).

## Full Example: 3-Slide Deck From Scratch

```
create_presentation(filePath: "C:\Decks\q4.pptx")
open_presentation(filePath: "C:\Decks\q4.pptx") → sessionId

# Slide 1: title
add_slide(sessionId) → slideIndex=1
set_layout(sessionId, 1, "ppLayoutTitle")
add_text_box(sessionId, 1, left=50, top=50, width=600, height=80, text="Q4 Results")
set_font_size(sessionId, 1, 1, fontSize=36)
set_bold(sessionId, 1, 1, bold=true)
set_notes_text(sessionId, 1, "Welcome the audience and set the scope for Q4 review.")

# Slide 2: chart
add_slide(sessionId) → slideIndex=2
add_chart(sessionId, 2, "bar", left=50, top=100, width=500, height=300,
          categories=["Q1","Q2","Q3","Q4"], seriesName="Revenue", values=[120,150,170,210])
set_notes_text(sessionId, 2, "Revenue grew steadily each quarter, accelerating in Q4.")

# Slide 3: table
add_slide(sessionId) → slideIndex=3
add_table(sessionId, 3, rows=3, columns=2, left=50, top=100, width=400, height=200)
set_cell_text(sessionId, 3, 1, 1, 1, "Region")
set_cell_text(sessionId, 3, 1, 1, 2, "Growth")
# ... remaining cells ...

# Verify
export_all_slides_to_images(sessionId, outputDirectory: "C:\Decks\preview")
# Look at Slide1.PNG, Slide2.PNG, Slide3.PNG

save_presentation(sessionId)
close_presentation(sessionId)
```
