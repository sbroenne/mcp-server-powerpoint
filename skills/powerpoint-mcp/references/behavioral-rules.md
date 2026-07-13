# Behavioral Rules for PowerPoint MCP Operations

These rules ensure efficient, reliable PowerPoint automation via a live PowerPoint desktop
instance (COM). AI assistants should follow these guidelines when using the **13 PowerPoint MCP
tools across 13 domains**.

## Core Execution Rules

- **Execute tasks immediately without asking for confirmation.** Make reasonable assumptions
  (slide count, positions, colors) and proceed.
- **Never ask clarifying questions for standard operations.** Use `presentation(action: "list")`
  to discover open sessions, `slide(action: "get-count", session_id: ...)` to discover slide
  range, `shape(action: "get-count", session_id: ..., slide_index: ...)` to discover shapes on a
  slide — do not ask the user for information you can look up yourself.
- **Always end with a text summary.** Never end a turn with only a tool call. After finishing,
  state what was created/changed, the file path, and the slide count.

## Session Model (CRITICAL)

Every editing workflow starts by establishing a session:

```
1. presentation(action: "create", filePath: ...) OR presentation(action: "open", filePath: ...) → returns sessionId
2. ... all other domain tools take session_id; presentation lifecycle/property actions take sessionId ...
3. presentation(action: "save", sessionId: ...)   → persists changes to disk
4. presentation(action: "close", sessionId: ...)  → releases the session; PowerPoint shuts down in background
```

- `presentation(action: "create", ...)` creates a new file **and leaves the session open**. Do
  **not** follow it with a second open call on the same file unless you intentionally want another
  session.
- `sessionId` is opaque — do not try to construct or guess one. Always use the value returned by
  `presentation(action: "create"/"open", ...)`.
- Unknown/expired `sessionId` values return `success: false` with `errorMessage: "Unknown
  sessionId: ..."` — reopen the file to get a fresh session, do not retry the same id.
- `presentation(action: "list")` shows every open session (`sessionId`, `presentationPath`,
  `isPowerPointProcessAlive`) — use it to check state instead of asking the user which
  presentation is open.

## Tool Conventions

- **All 13 MCP tools are action-dispatch tools.** Every call includes an `action` parameter.
- **`presentation` uses camelCase lifecycle/property parameters** — `filePath`, `sessionId`,
  `templatePath`, `propertyName`, `value`.
- **The other 12 domain tools use `session_id` plus snake_case action parameters**, e.g.
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
`errorMessage` and correct the index instead of blindly retrying.

## Explicit Save Is Required

Domain tool actions (`slide(action: "add-blank", ...)`, `textframe(action: "set-text", ...)`,
`chart(action: "add-chart", ...)`, etc.) modify the **in-memory** presentation only. Nothing is
written to disk until you call `presentation(action: "save", sessionId: ...)`. If you close a
session without saving first, all changes since the last save are lost.

```
1. slide(action: "add-blank", session_id: ...)                                      → slide added in memory
2. textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ...) → text set in memory
3. presentation(action: "save", sessionId: ...)                                     → NOW persisted to disk
4. presentation(action: "close", sessionId: ...)                                    → safe to close
```

## Close Is Asynchronous (Do NOT Wait For It)

`presentation(action: "close", sessionId: ...)` returns as soon as the session is removed from the
registry — it **does not** wait for the underlying PowerPoint process to fully exit. Office's own
post-Quit cleanup can legitimately take up to a few minutes; this is normal COM/Office behavior,
not a hung call or a leaked process.

- Do not poll `presentation(action: "list")` waiting for the process to disappear — the session
  itself is already gone from the list immediately.
- Do not treat a slow-to-exit `POWERPNT.exe` in Task Manager as a bug.
- If you need to open the same file again immediately after closing it, a brief delay may be
  needed for the OS file lock to clear.

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
