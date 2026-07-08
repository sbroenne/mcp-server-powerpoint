---
name: powerpoint-cli
description: >
  PowerPoint CLI automation skill for Windows presentations. Use when a coding agent needs
  token-efficient, scriptable, or unattended PowerPoint automation via pptcli commands. Best
  for CI/CD, scheduled jobs, batch processing, PowerShell workflows, and bulk deck edits.
  Supports slides, shapes, text frames, tables, charts, images, speaker notes, layouts, and
  .potx template restyling. Triggers: pptcli, PowerPoint CLI, command line, batch, script,
  automation, CI/CD, scheduled, PowerShell, unattended, coding agent, deck processing.
---

# PowerPoint Automation with pptcli

## Preconditions

- Windows host with Microsoft PowerPoint installed (2016+)
- Uses COM interop — does NOT work on macOS or Linux
- Requires `pptcli.exe` on PATH (`dotnet tool install --global Sbroenne.PowerPointMcp.CLI`)

## How pptcli Works — a Background Daemon, Not a Cold Start Per Command

Launching PowerPoint and tearing it down again can take 90-150 seconds. To avoid paying that
cost on every single CLI invocation, `pptcli` keeps one PowerPoint instance alive in a small
background daemon (`pptcli service run`), auto-started the first time you open or create a
session, and reused by every subsequent command that references the same session id — even
though each command is a separate OS process.

## Workflow Checklist

| Step | Command | When |
|------|---------|------|
| 1. Session | `session create/open` | Always first |
| 2. Slides | `slide add-blank` | If needed |
| 3. Edit content | See command reference below | Shapes, text, tables, charts, images, notes |
| 4. Save & close | `session close --save` | Always last |

## CRITICAL RULES (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

Execute commands to discover the answer instead:

| DON'T ASK | DO THIS INSTEAD |
|-----------|-----------------|
| "Which file should I use?" | `pptcli session list` |
| "How many slides does the deck have?" | `pptcli slide get-count -s <id>` |
| "How many shapes are on this slide?" | `pptcli shape get-count -s <id> --slide-index N` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a command execution.** After completing all operations, always
provide a brief text message confirming what was done. Silent command-only responses are
incomplete.

### Rule 3: Session Lifecycle

**Creating vs Opening Files:**
```powershell
# NEW file - use session create
pptcli session create C:\decks\demo.pptx   # Creates file + returns session ID

# EXISTING file - use session open
pptcli session open C:\decks\demo.pptx     # Opens file + returns session ID
```

**CRITICAL: Use `session create` for new files. `session open` on non-existent files will fail!**

**CRITICAL: ALWAYS use the session ID returned by `session create` or `session open` in
subsequent commands via `-s`/`--session`. NEVER guess or hardcode session IDs. The session ID is
in the JSON output (e.g., `{"sessionId":"abc123"}`). Parse it and use it.**

```powershell
# Example: capture session ID from output, then use it
pptcli session create C:\decks\demo.pptx        # Returns JSON with sessionId
pptcli slide add-blank -s <returned-session-id>
pptcli session close <returned-session-id> --save
```

**Unclosed sessions leave PowerPoint processes running, locking files** — but the daemon shuts
down automatically after an idle period with no open sessions, or immediately via
`pptcli service stop`.

### Rule 4: 1-Based Indexing Everywhere

`--slide-index`, `--shape-index`, `--row`, `--column` are all **1-based**, matching PowerPoint's
native object model — not 0-based.

### Rule 5: Restyling an Existing Deck via a .potx Template

```powershell
$session = pptcli session open C:\decks\demo.pptx | ConvertFrom-Json
pptcli presentation apply-template -s $session.sessionId --template-path C:\templates\corp.potx
pptcli presentation get-theme-name -s $session.sessionId   # verify the theme actually changed
pptcli session close $session.sessionId --save
```

Slide content is preserved; only masters/theme/layouts change.

### Rule 6: Report File Errors Immediately

If you see "File not found" or "Path not found" - STOP and report to user. Don't retry.

### Rule 7: Managing the Daemon Directly

```powershell
pptcli service status         # Report whether the daemon is running, session count, uptime
pptcli service stop           # Graceful shutdown
pptcli service stop --force   # Force-kill if a graceful shutdown doesn't respond
```

## CLI Command Reference

**Syntax rule:** CLI commands use `pptcli <command> <action> -s <SESSION_ID> --kebab-case-flags ...`.
Do not use MCP call syntax such as `shape(action: "add-rectangle", session_id: ..., slide_index: ...)` or
snake_case parameters — the CLI uses kebab-case flags and a `<command> <action>` shape instead
of one flat tool per verb.

Available command groups (in addition to `session` and `service`):

`chart`, `image`, `layout`, `notes`, `presentation`, `shape`, `slide`, `table`, `textframe`

Run `pptcli <command> --help` for the live, authoritative list of actions and flags for that
command — the table below is a summary generated from the same Core interfaces as the MCP tool
surface, so it can lag a `--help` run for in-flight changes. See
[`references/cli-commands.md`](./references/cli-commands.md) for the full command reference
(all domains, actions, and parameters in one document).


### `chart` — Chart lifecycle and data operations.

Actions: `add-chart`, `get-chart-data`

| Flag | Description |
|------|-------------|
| `--slide-index` | 1-based slide index. (required) |
| `--chart-type` | Chart type: "bar", "line", or "pie". (required for: add-chart) |
| `--left` | Left position in points. (required for: add-chart) |
| `--top` | Top position in points. (required for: add-chart) |
| `--width` | Width in points. (required for: add-chart) |
| `--height` | Height in points. (required for: add-chart) |
| `--categories` | Category labels (x-axis / pie slice labels). (required for: add-chart) |
| `--series-name` | Name of the single data series. (required for: add-chart) |
| `--values` | Data values, one per category. (required for: add-chart) |
| `--shape-index` | (required for: get-chart-data) |


### `image` — Image commands: embed a picture file into a slide. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

Actions: `add-picture`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--image-path` | (required) |
| `--left` | (required) |
| `--top` | (required) |
| `--width` | (required) |
| `--height` | (required) |


### `layout` — Slide layout commands: apply/read a slide's built-in layout. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

Actions: `set-layout`, `get-layout`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--layout-name` | (required for: set-layout) |


### `notes` — Speaker notes commands: set/get the notes text for a slide. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

Actions: `set-notes-text`, `get-notes-text`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--text` | (required for: set-notes-text) |


### `presentation` — Presentation lifecycle commands: create, close, save.

Actions: `create`, `open`, `save`, `apply-template`, `get-theme-name`

| Flag | Description |
|------|-------------|
| `--file-path` | (required for: create, open) |
| `--is-macro-enabled` |  |
| `--template-path` | Full path to a .potx/.potm/.pot template file (a .pptx/.pptm presentation may also be used as a template source, matching PowerPoint's own behavior). (required for: apply-template) |


### `shape` — Shape commands: add rectangles/text boxes, count, delete, reposition/resize. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

Actions: `add-rectangle`, `add-text-box`, `get-count`, `delete`, `set-position`, `set-size`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--left` | (required for: add-rectangle, add-text-box, set-position) |
| `--top` | (required for: add-rectangle, add-text-box, set-position) |
| `--width` | (required for: add-rectangle, add-text-box, set-size) |
| `--height` | (required for: add-rectangle, add-text-box, set-size) |
| `--text` | (required for: add-text-box) |
| `--shape-index` | (required for: delete, set-position, set-size) |


### `slide` — Slide lifecycle commands: add, delete, count. First domain built on top of the presentation lifecycle commands, operating within an already-open IPresentationBatch.

Actions: `add-blank`, `get-count`, `delete`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required for: delete) |


### `table` — Table commands: add a table shape and read/write cell text. Operates within an already-open IPresentationBatch, targeting a specific slide and table shape by their 1-based indices.

Actions: `add-table`, `set-cell-text`, `get-cell-text`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--rows` | (required for: add-table) |
| `--columns` | (required for: add-table) |
| `--left` | (required for: add-table) |
| `--top` | (required for: add-table) |
| `--width` | (required for: add-table) |
| `--height` | (required for: add-table) |
| `--shape-index` | (required for: set-cell-text, get-cell-text) |
| `--row` | (required for: set-cell-text, get-cell-text) |
| `--column` | (required for: set-cell-text, get-cell-text) |
| `--text` | (required for: set-cell-text) |


### `textframe` — Text frame commands: set/get text and basic font formatting (size, bold, color) for a shape's text range. Operates within an already-open IPresentationBatch, targeting a specific shape by its 1-based slide and shape index.

Actions: `set-text`, `get-text`, `set-font-size`, `set-bold`, `set-font-color`

| Flag | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required) |
| `--text` | (required for: set-text) |
| `--font-size` | (required for: set-font-size) |
| `--bold` | (required for: set-bold) |
| `--red` | (required for: set-font-color) |
| `--green` | (required for: set-font-color) |
| `--blue` | (required for: set-font-color) |



## Common Pitfalls

- `--red`/`--green`/`--blue` for `textframe set-font-color` are each 0-255 integers, not a single
  hex string.
- A session created with `session create` is already open — do not call `session open` on the
  same path again; that opens a SECOND session against the same file.
