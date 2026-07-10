# Behavioral Rules for PowerPoint MCP Operations

These rules ensure efficient, reliable PowerPoint automation via a live PowerPoint desktop
instance (COM). AI assistants should follow these guidelines when using the 18 PowerPoint MCP
tools (7 session-lifecycle tools + 11 domain action-dispatch tools).

## Core Execution Rules

- **Execute tasks immediately without asking for confirmation.** Make reasonable assumptions
  (slide count, positions, colors) and proceed.
- **Never ask clarifying questions for standard operations.** Use `list_sessions` to discover
  open sessions, `slide(action: "get-count", session_id: ...)` to discover slide range,
  `shape(action: "get-count", session_id: ..., slide_index: ...)` to discover shapes on a slide —
  don't ask the user for information you can look up yourself.
- **Always end with a text summary.** Never end a turn with only a tool call. After finishing,
  state what was created/changed, the file path, and the slide count.

## Session Model (CRITICAL)

Every editing operation requires an open session:

```
1. open_presentation(filePath)        → returns sessionId
2. ... all domain tools take session_id ...
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

## Two Tool Families, Two Calling Conventions

- **Session-lifecycle tools** (`create_presentation`, `open_presentation`, `save_presentation`,
  `close_presentation`, `list_sessions`, `apply_template`, `get_theme_name`,
  `set_document_property`, `get_document_property`, `set_custom_property`, `get_custom_property`,
  `remove_custom_property`) are one-tool-per-verb and use camelCase arguments, e.g.
  `open_presentation(filePath: "C:\\Decks\\q4.pptx")`.
- **Domain tools** (`slide`, `shape`, `textframe`, `table`, `chart`, `image`, `notes`, `layout`,
  `master`, `smartart`, `animation`, `export`) are action-dispatch: one tool per domain, and every call passes an `action` (kebab-case)
  plus `session_id` and only the snake_case parameters relevant to that action, e.g.
  `shape(action: "add-rectangle", session_id: ..., slide_index: 1, left: 50, top: 80, width: 100,
  height: 60)`.

## 1-Based Indexing (CRITICAL — the #1 source of bugs)

**Every index in the PowerPoint MCP surface is 1-based, matching PowerPoint's own object model**
(`Slides(1)` is the first slide, not `Slides(0)`):

- `slide_index` — 1 is the first slide.
- `shape_index` — 1 is the first shape added to a slide.
- Table `row` / `column` — 1 is the first row/column.

This differs from most programming languages (0-based arrays) and from some other Office MCP
servers. Passing `0` or a negative index returns `success: false`, never an exception — check the
`errorMessage` and correct the index, don't blindly retry.

## Explicit Save Is Required

Domain tool actions (`slide(action: "add-blank", ...)`, `textframe(action: "set-text", ...)`,
`chart(action: "add-chart", ...)`, etc.) modify the **in-memory** presentation only. Nothing is
written to disk until you call `save_presentation(sessionId)`. If you close a session without
saving first, all changes since the last save are lost.

```
1. slide(action: "add-blank", session_id: ...)                                    → slide added in memory
2. textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ..., text: ...) → text set in memory
3. save_presentation(sessionId)                                                    → NOW persisted to disk
4. close_presentation(sessionId)                                                   → safe to close, nothing pending
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
1. shape(action: "add-rectangle", ...) / textframe(action: "set-text", ...) / chart(action: "add-chart", ...)  → make the change
2. export(action: "export-slide-to-image", session_id: ..., slide_index: ..., output_path: ...)                → render it
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
