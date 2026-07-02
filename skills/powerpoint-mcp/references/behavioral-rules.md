# Behavioral Rules for PowerPoint MCP Operations

These rules ensure efficient, reliable PowerPoint automation via a live PowerPoint desktop
instance (COM). AI assistants should follow these guidelines when using the 31 PowerPoint MCP
tools.

## Core Execution Rules

- **Execute tasks immediately without asking for confirmation.** Make reasonable assumptions
  (slide count, positions, colors) and proceed.
- **Never ask clarifying questions for standard operations.** Use `list_sessions` to discover
  open sessions, `get_slide_count` to discover slide range, `get_shape_count` to discover shapes
  on a slide — don't ask the user for information you can look up yourself.
- **Always end with a text summary.** Never end a turn with only a tool call. After finishing,
  state what was created/changed, the file path, and the slide count.

## Session Model (CRITICAL)

Every editing operation requires an open session:

```
1. open_presentation(filePath)        → returns sessionId
2. ... all other tools take sessionId  ...
3. save_presentation(sessionId)        → persists changes to disk
4. close_presentation(sessionId)       → releases the PowerPoint process
```

- `create_presentation` creates a new file **on disk and does NOT open a session** — it calls
  Create, saves, and closes internally. You MUST call `open_presentation` afterward to add
  slides/content to a freshly created file. This is a deliberate two-call flow, not a bug.
- `sessionId` is opaque — do not try to construct or guess one. Always use the value returned by
  `open_presentation`.
- Unknown/expired `sessionId` values return `success: false` with `errorMessage: "Unknown
  sessionId: ..."` — re-open the file to get a fresh session, don't retry the same id.
- `list_sessions` shows every open session (`sessionId`, `presentationPath`,
  `isPowerPointProcessAlive`) — use it to check state instead of asking the user "which
  presentation?".

## 1-Based Indexing (CRITICAL — the #1 source of bugs)

**Every index in the PowerPoint MCP surface is 1-based, matching PowerPoint's own object model**
(`Slides(1)` is the first slide, not `Slides(0)`):

- `slideIndex` — 1 is the first slide.
- `shapeIndex` — 1 is the first shape added to a slide.
- Table `row` / `column` — 1 is the first row/column.

This differs from most programming languages (0-based arrays) and from some other Office MCP
servers. Passing `0` or a negative index returns `success: false`, never an exception — check the
`errorMessage` and correct the index, don't blindly retry.

## Explicit Save Is Required

Operations (`add_slide`, `set_text`, `add_chart`, etc.) modify the **in-memory** presentation
only. Nothing is written to disk until you call `save_presentation(sessionId)`. If you close a
session without saving first, all changes since the last save are lost.

```
1. add_slide(sessionId) → slide added in memory
2. set_text(sessionId, slideIndex, shapeIndex, text) → text set in memory
3. save_presentation(sessionId) → NOW persisted to disk
4. close_presentation(sessionId) → safe to close, nothing pending
```

## Close Is Asynchronous (Do NOT Wait For It)

`close_presentation` returns as soon as the session is removed from the registry — it does
**not** wait for the underlying PowerPoint process to fully exit. Office's own post-Quit cleanup
can legitimately take up to a few minutes; this is normal COM/Office behavior, not a hung call or
a leaked process.

- Do NOT poll `list_sessions` waiting for the process to disappear — the session itself is
  already gone from the list immediately.
- Do NOT treat a slow-to-exit `POWERPNT.exe` in Task Manager as a bug.
- If you need to open the SAME file again immediately after closing it, a brief delay may be
  needed for the OS file lock to clear; prefer opening a different file or waiting briefly if you
  see a file-lock error.

## Verify Visually (Our Differentiator)

Text-only inspection cannot catch overlapping shapes, overflowing text, or bad chart layouts.
After creating or changing visual content, export and look at the result:

```
1. add_shape / set_text / add_chart / ...  → make the change
2. export_slide_to_image(sessionId, slideIndex, outputPath) → render it
3. Look at the returned image → confirm it matches intent, fix if not
```

See `export-and-verify.md` for the full loop and when it is required.

## Report Results

After completing operations, report:

- What was created/modified (slide count, shapes added, text set).
- The file path.
- Whether the presentation was saved.

**Bad:** *(tool call with no text)*
**Good:** "Added 3 slides with title + content layout to `C:\Decks\q4.pptx`, exported slide 1 for
review, and saved the file."
