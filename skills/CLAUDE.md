# CLAUDE.md — PowerPoint MCP Server Skills

Project instructions for Claude Code (and Claude Desktop via the `powerpoint-mcp` Agent Skill)
when working with the PowerPoint MCP Server.

## What This Is

`mcp-server-powerpoint` drives a live Microsoft PowerPoint desktop instance via COM
(`Microsoft.Office.Interop.PowerPoint`) and exposes it as **13 MCP tools across 13 domains**.
Every domain is a single action-dispatch tool: `presentation`, `slide`, `shape`, `textframe`,
`table`, `notes`, `layout`, `master`, `animation`, `image`, `chart`, `smartart`, and `export`.

Windows + PowerPoint desktop required. There is no cross-platform or headless mode — everything
goes through real COM automation of a real PowerPoint process.

## Skill Source of Truth

`skills/shared/*.md` is the single source of truth for all authoring guidance. It is copied into
`skills/powerpoint-mcp/references/` for skill-based clients (Claude Code, Cursor, Windsurf, VS
Code) and is intended to also back auto-generated `[McpServerPrompt]` methods once that tooling
exists (see `.github/copilot-instructions.md`).

**Never edit `skills/powerpoint-mcp/references/*.md` directly** — edit the corresponding file in
`skills/shared/` and re-copy. Editing the copy only will be silently overwritten by the next sync.

## Core Behavioral Rules (see `skills/shared/behavioral-rules.md` for full detail)

1. **Sessions required** — start with `presentation(action: "create", filePath: ...)` or
   `presentation(action: "open", filePath: ...)`, then reuse the returned `sessionId` for all edit
   calls.
2. **1-based indexing everywhere** — `slide_index`, `shape_index`, table `row`/`column` all start
   at 1.
3. **Explicit save** — nothing persists to disk until `presentation(action: "save", sessionId: ...)`.
4. **Close is async** — `presentation(action: "close", sessionId: ...)` does not wait for the
   PowerPoint process to exit.
5. **Verify visually** — use `export(action: "export-slide-to-image"/"export-all-slides-to-images",
   ...)` after any visual change; this is the project's core differentiator over text-only
   PowerPoint tooling.
6. **Never ask clarifying questions** — discover state with `presentation(action: "list")`,
   `slide(action: "get-count", ...)`, `shape(action: "get-count", ...)`, etc.
7. **Always end with a text summary** — never end a turn on a bare tool call.

## When Building Decks

Follow `skills/shared/workflows.md` and `skills/shared/deck-builder.md`: plan the full slide
order up front, create/open one session, build slide-by-slide with the create → verify → fix loop,
and vary layout composition across slides.

## Repository Conventions (for contributing to this repo itself)

See `.github/copilot-instructions.md` for the C# architecture (ComInterop → Core → McpServer/CLI),
Rule 1/1b (Success/ErrorMessage invariant), and Rule 30 (real-COM integration-test-only TDD). This
file is about the PowerPoint *skill content*, not the C# codebase.
