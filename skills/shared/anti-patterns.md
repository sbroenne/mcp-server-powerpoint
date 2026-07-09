# Anti-Patterns to Avoid

Common mistakes when using the PowerPoint MCP tools. These cause errors, data loss, or wasted
turns — avoid them.

## Forgetting to Open a Session

### The Problem

Calling a domain tool without a valid `session_id`:

```
WRONG:
slide(action: "add-blank", session_id: "made-up-id")  → success: false, "Unknown sessionId"
```

### The Solution

Always open a session first and reuse the exact `sessionId` string it returns:

```
CORRECT:
result = open_presentation(filePath: "C:\Decks\q4.pptx")
sessionId = result.sessionId
slide(action: "add-blank", session_id: sessionId)
```

## Wrong Index Base (0-Based Instead of 1-Based)

### The Problem

Treating `slide_index`/`shape_index`/table `row`/`column` as 0-based, matching most programming
languages:

```
WRONG: assuming the first slide is index 0
shape(action: "get-count", session_id: ..., slide_index: 0)  → success: false (out of range)
```

### The Solution

Every index in this tool surface is 1-based, matching PowerPoint's own object model:

```
CORRECT:
shape(action: "get-count", session_id: ..., slide_index: 1)  → the first slide
```

See `behavioral-rules.md` for the full indexing rule.

## Forgetting to Save

### The Problem

Making changes, then closing without saving:

```
WRONG:
slide(action: "add-blank", session_id: ...)
textframe(action: "set-text", session_id: ..., ...)
close_presentation(sessionId)   → changes since last save are LOST
```

### The Solution

Always call `save_presentation` before `close_presentation` when changes were made:

```
CORRECT:
slide(action: "add-blank", session_id: ...)
textframe(action: "set-text", session_id: ..., ...)
save_presentation(sessionId)
close_presentation(sessionId)
```

## Expecting `close_presentation` to Block

### The Problem

Waiting for or polling after `close_presentation`, assuming it doesn't return until PowerPoint's
process has fully exited:

```
WRONG:
close_presentation(sessionId)
list_sessions()  → repeatedly poll, waiting for POWERPNT.exe to disappear from Task Manager
```

### The Solution

`close_presentation` returns as soon as the session is removed from the registry; PowerPoint's own
process cleanup happens in the background afterward and can take up to a few minutes — this is
normal Office behavior, not a bug. The session is already gone from `list_sessions` immediately;
don't wait on the OS process.

## Skipping Visual Verification

### The Problem

Trusting `success: true` from a shape/chart/table/image call as proof the slide looks right:

```
WRONG:
chart(action: "add-chart", session_id: ..., slide_index: ..., ...)  → success: true
save_presentation(sessionId)
close_presentation(sessionId)
# Never looked at the rendered slide — chart could be mis-sized, overlapping, or have wrong data
```

### The Solution

Export and look at the result before saving/closing when visual content was added or changed (see
`export-and-verify.md`):

```
CORRECT:
chart(action: "add-chart", session_id: ..., slide_index: ..., ...)
export(action: "export-slide-to-image", session_id: ..., slide_index: ..., output_path: ...)
# Inspect the image, fix issues found
save_presentation(sessionId)
close_presentation(sessionId)
```

## Delete-and-Rebuild for Small Changes

### The Problem

Deleting and re-creating a shape/table/chart to make a small change:

```
WRONG: fixing one table cell
shape(action: "delete", session_id: ..., slide_index: ..., shape_index: ...)
table(action: "add-table", session_id: ..., slide_index: ..., rows: 4, columns: 3, ...)
# ... re-populate every cell from scratch ...
```

### The Solution

Use the targeted update action for the specific thing that changed:

```
CORRECT:
table(action: "set-cell-text", session_id: ..., slide_index: ..., shape_index: ..., row: 3, column: 2, text: "$1.8M")
```

Same principle for shapes: prefer `shape(action: "set-position", ...)`/`shape(action: "set-size",
...)` over delete-and-recreate (see `slides-and-shapes.md`).

## Session Leaks

### The Problem

Opening several presentations and never closing them:

```
WRONG:
open_presentation("file1.pptx")  → session 1
open_presentation("file2.pptx")  → session 2
open_presentation("file3.pptx")  → session 3
# ... none ever closed
```

Each open session holds its own `POWERPNT.exe` process. Left unclosed, they accumulate, consume
memory, and keep file locks on disk.

### The Solution

Close each session when its work is done, saving first if changes were made:

```
CORRECT:
s1 = open_presentation("file1.pptx")
# ... work ...
save_presentation(s1)
close_presentation(s1)

s2 = open_presentation("file2.pptx")
# ... work ...
save_presentation(s2)
close_presentation(s2)
```

## Assuming Multi-Series Charts

### The Problem

Trying to pass multiple series into a single `chart(action: "add-chart", ...)` call — the action
only accepts one `series_name` + one `values` array (see `charts.md`).

### The Solution

Pick the single most important series, or place multiple `chart(action: "add-chart", ...)` calls
side-by-side on the slide with separate labels, rather than expecting a multi-series parameter
that doesn't exist in this surface.

## Guessing Layout Names

### The Problem

Passing an invented or approximate `layout_name` to `layout(action: "set-layout", ...)` (e.g.,
`"Blank"`, `"TitleSlide"`):

```
WRONG:
layout(action: "set-layout", session_id: ..., slide_index: ..., layout_name: "TitleSlide")  → success: false
```

### The Solution

Use the exact `PpSlideLayout` enum member name (`ppLayoutTitle`, `ppLayoutBlank`, etc.) — see
`layouts.md` for the common ones.

## Discovery Loop Without Action

### The Problem

Repeating `list_sessions`, `slide(action: "get-count", ...)`, or `shape(action: "get-count", ...)`
multiple times without acting on the result:

```
WRONG:
shape(action: "get-count", session_id: ..., slide_index: ...)  → 3
shape(action: "get-count", session_id: ..., slide_index: ...)  → 3 (same call again)
list_sessions()                                                → unrelated re-check
shape(action: "get-count", session_id: ..., slide_index: ...)  → 3 (again)
```

### The Solution

Call a discovery action once, then act on what it returned. If a session genuinely expired,
re-open it once and continue — don't loop on the same discovery call more than twice in a row
(see `behavioral-rules.md`).
