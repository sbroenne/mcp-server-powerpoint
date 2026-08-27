> **CLI syntax:** Shared guides may use MCP calls as shorthand. Use `cli-commands.md` or live `--help` for exact commands and kebab-case options.

# Behavioral Rules for PowerPoint MCP Operations

These rules ensure efficient, reliable PowerPoint automation via a live PowerPoint desktop
instance (COM). AI assistants should follow these guidelines when using the **16 PowerPoint MCP
tools across 16 domains**.

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
3. presentation(action: "close", sessionId: ..., save: true) → persists changes and releases the session
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

- **All 16 MCP tools are action-dispatch tools.** Every call includes an `action` parameter.
- **`presentation` uses camelCase lifecycle/property parameters** — `filePath`, `sessionId`,
  `targetPath`, `format`, `overwrite`, `templatePath`, `propertyName`, `value`.
- **The other 15 domain tools use `session_id` plus snake_case action parameters**, e.g.
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

## Explicit Save-on-Close Is Required

Domain tool actions (`slide(action: "add-blank", ...)`, `textframe(action: "set-text", ...)`,
`chart(action: "add-chart", ...)`, etc.) modify the **in-memory** presentation only. Nothing is
written to disk unless you close with `save: true`. Closing with the default `save: false` discards
all changes since the last save.

```
1. slide(action: "add-blank", session_id: ...)                                      → slide added in memory
2. textframe(action: "set-text", session_id: ..., slide_index: ..., shape_index: ...) → text set in memory
3. presentation(action: "close", sessionId: ..., save: true)                         → persisted and closed
```

`save-as` and `save-copy-as` are explicit delivery operations, not a generic save action:

- `save-as` writes `.pptx`, `.pptm`, or `.ppt`, then changes the active session path only after
  PowerPoint succeeds.
- `save-copy-as` requires the destination extension to match the active presentation and leaves
  the active session path unchanged.
- Both reject an existing destination unless `overwrite: true` is supplied.

## Mark as Final Is Advisory, Not Security

Use `presentation(action: "get-final", sessionId: ...)` to read PowerPoint's Mark as Final state
and `presentation(action: "set-final", sessionId: ..., isFinal: true/false)` to set or clear it.
This flag only communicates that editing is discouraged. It is not authentication, encryption, or
access control, and anyone can clear it.

Setting the flag to `true` first saves all current changes, then PowerPoint persists the flag and
makes the presentation read-only. Calling
`presentation(action: "close", sessionId: ..., save: true)` remains valid and closes the session
without attempting a forbidden second save, so edits made before `set-final` are not lost. After
clearing the flag with `isFinal: false`, close with `save: true` to persist the cleared state.

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

## Run the Deterministic Accessibility Audit

Before final delivery, call `accessibility(action: "audit", session_id: ...)`. Fix missing
alternative text and empty title placeholders, then rerun the audit. This is a deterministic
PowerPoint structure check, not an AI review of writing quality.

## Report Results

After completing operations, report:

- What was created/modified (slide count, shapes added, text set).
- The file path.
- Whether the presentation was saved.

**Bad:** *(tool call with no text)*
**Good:** "Added 3 slides with title + content layout to `C:\Decks\q4.pptx`, exported slide 1 for
review, and saved the file."
