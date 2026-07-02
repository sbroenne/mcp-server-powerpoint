# Copilot Coding Agent — Squad Instructions

You are working on a project that uses **Squad**, an AI team framework. When picking up issues autonomously, follow these guidelines.

## Coordinator Canary Check

**IMPORTANT — run this check on EVERY session start, before doing any work:**

Look for the exact token `SQUAD_COORDINATOR_CANARY_a8f3` in your `agent_instructions`. This token is placed at the end of `squad.agent.md`. If it is missing, the coordinator file is missing or truncated.

**When the token is missing, you MUST:**
1. **STOP** — do not proceed with standard Squad behavior.
2. **WARN the user** with this exact message:
   ```
   ⚠️ Squad coordinator (squad.agent.md) appears to be missing or truncated. The canary token was not found. Do NOT proceed with standard Squad behavior — Squad's safety rails are not loaded. Please restart your session.
   ```
3. Do not continue with normal Squad routing, spawning, PR, or branch-protection behavior after emitting the warning.

## Team Context

Before starting work on any issue:

1. Read `.squad/team.md` for the team roster, member roles, and your capability profile.
2. Read `.squad/routing.md` for work routing rules.
3. If the issue has a `squad:{member}` label, read that member's charter at `.squad/agents/{member}/charter.md` to understand their domain expertise and coding style — work in their voice.

## Capability Self-Check

Before starting work, check your capability profile in `.squad/team.md` under the **Coding Agent → Capabilities** section.

- **🟢 Good fit** — proceed autonomously.
- **🟡 Needs review** — proceed, but note in the PR description that a squad member should review.
- **🔴 Not suitable** — do NOT start work. Instead, comment on the issue:
  ```
  🤖 This issue doesn't match my capability profile (reason: {why}). Suggesting reassignment to a squad member.
  ```

## Branch Naming

Use the squad branch convention:
```
squad/{issue-number}-{kebab-case-slug}
```
Example: `squad/42-fix-login-validation`

## PR Guidelines

When opening a PR:
- Reference the issue: `Closes #{issue-number}`
- If the issue had a `squad:{member}` label, mention the member: `Working as {member} ({role})`
- If this is a 🟡 needs-review task, add to the PR description: `⚠️ This task was flagged as "needs review" — please have a squad member review before merging.`
- Follow any project conventions in `.squad/decisions.md`

## Decisions

If you make a decision that affects other team members, write it to:
```
.squad/decisions/inbox/copilot-{brief-slug}.md
```
The Scribe will merge it into the shared decisions file.

---

# Project Instructions — PowerPointMcp

> The Squad framework sections above are authoritative for coordination/routing. Everything below
> is project-specific technical guidance for this codebase.

## Critical Files (Read These First)

**ALWAYS read when working on code:**
- [Critical Rules](instructions/critical-rules.instructions.md) - mandatory rules (Success flag, COM cleanup, TDD, session lifecycle)
- [Architecture Patterns](instructions/architecture-patterns.instructions.md) - layered architecture, command pattern, resource management

**Read based on task type:**
- Adding/fixing Core commands or COM interop → [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- Writing tests → [Testing Strategy](instructions/testing-strategy.instructions.md)
- MCP Server tool work → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)

## What is PowerPointMcp?

**PowerPointMcp** is a Windows-only toolset for programmatic PowerPoint automation via COM
interop, designed for coding agents and automation scripts. It drives a **live PowerPoint desktop
instance** through the official `Microsoft.Office.Interop.PowerPoint` PIA (embedded via
`ForceEmbedPowerPointInteropTypes` — no runtime assembly-resolver needed, unlike Excel).

> **NOTE: Unlike `mcp-server-excel`, the CLI here is currently a hand-written placeholder, not an
> equal first-class entry point.** The MCP Server is the primary, actively-developed surface.
> Do not assume CLI/MCP parity exists yet — see `src/PowerPointMcp.CLI/Program.cs`.

**Core Layers:**
1. **ComInterop** (`src/PowerPointMcp.ComInterop`) - STA thread + OLE message filter + channel-based
   work queue (`PresentationBatch`), ported from `mcp-server-excel`'s `ExcelBatch` pattern.
2. **Core** (`src/PowerPointMcp.Core`) - PowerPoint domain commands, one folder per domain
   (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart, Export).
3. **McpServer** (`src/PowerPointMcp.McpServer`) - Model Context Protocol stdio host. Hand-written
   `[McpServerToolType]` classes (one per domain, 31 tools total) — no source generators yet
   (deferred until the Core shape is proven stable; see architecture decision in
   `.squad/decisions.md`, 2026-07-01T11-14-34).
4. **CLI** (`src/PowerPointMcp.CLI`) - Minimal hand-written placeholder proving
   ComInterop → Core → CLI end to end. NOT the target Generators-based CLI.

**No out-of-process Service** (unlike Excel's `ExcelMcp.Service` + named-pipe `ServiceBridge`).
The MVP uses an in-process `PresentationSessionRegistry` singleton
(`ConcurrentDictionary<string sessionId, IPresentationBatch>`) inside the MCP host — the STA
thread + work queue already live inside `PresentationBatch`, and the MCP stdio host is itself
long-lived, so no RPC layer is needed for the MVP. Out-of-process hardening is deferred to Phase 2.

## Session Model (MCP Server)

```
open_presentation(filePath) → sessionId (registry.Open)
... all other tools take sessionId, resolve batch via registry.TryGet ...
save_presentation(sessionId) → Core Commands.Save(batch)
close_presentation(sessionId) → registry.Close(sessionId) — removes from registry immediately,
                                  disposes the batch (and its PowerPoint process) on a background
                                  task. Does NOT block the caller.
```

`create_presentation` is intentionally a **standalone** call: it creates + saves + closes a new
file on disk via Core directly, with NO session left open. Callers must `open_presentation`
afterward to edit it. This is a deliberate two-call flow (see `.squad/decisions.md`,
2026-07-01T12:00:00+02:00), not a bug — do not "fix" it into an implicit open.

Host shutdown (`Ctrl+C`, stdin EOF, or normal exit) MUST dispose every open batch via
`PresentationSessionRegistry.DisposeAll()` — two-layered: an `IHostedService` on `StopAsync`, plus
a backstop in `Main`'s `finally`. Never let a PowerPoint process leak past MCP host shutdown.

## 1-Based Indexing (CRITICAL)

`slideIndex`, `shapeIndex`, and table `row`/`column` are 1-based everywhere in Core and the MCP
surface, matching PowerPoint's native COM object model (`Slides(1)` is the first slide). Do not
introduce 0-based indexing anywhere in Core, the MCP tools, or tests.

## Rule 1 / 1b — Success/ErrorMessage Invariant (CRITICAL)

Core commands return `{Domain}OperationResult` with `Success`/`ErrorMessage`:

- **Rule 1:** `Success == true` implies `ErrorMessage == null`. Never both.
- **Rule 1b:** Never wrap `batch.Execute()` in a try-catch that swallows exceptions and returns an
  error result. Let exceptions propagate naturally out of Core — `batch.Execute()`'s
  `TaskCompletionSource` plumbing handles them. Only the MCP tool boundary
  (`PowerPointToolsBase.ExecuteToolAction`) catches unexpected exceptions, logs the HResult to
  stderr, and serializes a structured error — this is the ONLY place a catch-all belongs.

Expected, caller-correctable failures (bad index, missing file, unknown session) are validated
BEFORE calling Core and returned as `Success=false` payloads — never thrown as exceptions.

```csharp
// Core: let exceptions propagate
public ShapeOperationResult AddRectangle(IPresentationBatch batch, int slideIndex, ...)
{
    return batch.Execute((ctx, ct) => {
        // COM object access; exceptions bubble up naturally
    });
}

// MCP Tool: validate expected bad input first, don't throw
if (!registry.TryGet(sessionId, out var batch))
{
    return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
}
```

## Rule 30 — Real-COM Integration Tests Only (CRITICAL)

**NEVER write unit tests with mocked COM objects.** All Core tests are real-COM integration tests
against a real PowerPoint instance, following strict TDD (red → green): write a failing test
first, watch it fail, implement, watch it pass.

- Tests are serialized via `xunit.runner.json` (`maxParallelThreads: 1`) — concurrent PowerPoint
  launches are not supported.
- Tests are tagged `[Trait("Category", "Integration")]` + `[Trait("Feature", "<Domain>")]` (e.g.
  `Feature=Shape`, `Feature=Chart`) for surgical `dotnet test --filter` runs.
- MCP-layer tests (`tests/PowerPointMcp.McpServer.Tests`) use the SDK's in-memory pipe transport
  (`Program.ConfigureTestTransport`) for protocol-level assertions (tools/list, schema, error
  envelope) WITHOUT launching PowerPoint; only the session-lifecycle round-trip tests touch real
  COM. Do not re-test Core COM behavior at the MCP layer — MCP tests assert wiring + serialization
  + session mapping only.
- Office's own post-Quit cleanup can take up to ~90-200+ seconds after `close_presentation` /
  session disposal — this is documented, benign Office behavior. Do NOT add force-kill on the
  happy path.

### Test Commands

```powershell
# Full Core test suite (real COM, serialized — slower than a typical unit-test suite)
dotnet test tests\PowerPointMcp.Core.Tests

# Surgical: one domain only
dotnet test tests\PowerPointMcp.Core.Tests --filter "Feature=Shape"
dotnet test tests\PowerPointMcp.Core.Tests --filter "Feature=Chart"

# MCP transport/protocol tests
dotnet test tests\PowerPointMcp.McpServer.Tests
```

Always run tests with an explicit timeout in the terminal/tooling layer — never leave a COM test
run open-ended; fail fast if PowerPoint stalls.

## Quick Reference

### Tool Selection
- Code changes → targeted `edit`/`replace_string_in_file` (small, precise context)
- Find code → `grep`/code search over PowerShell text scraping
- Build/test/git → `dotnet build`/`dotnet test`/`git` via the terminal tool

### GitHub Auth Rule for This Repo

When using `gh` against `sbroenne/mcp-server-powerpoint` (issues, PRs, comments, merges), verify
`gh auth status` shows the personal `sbroenne` account as ACTIVE, not an EMU
(`stbrnner_microsoft`) account — EMU accounts cannot create/push to this namespace. Run
`gh auth switch --user sbroenne` if needed (see `.squad/decisions.md`).

## Key Lessons (Update After Major Work)

**Success Flag:** NEVER `Success = true` with `ErrorMessage`. Set `Success` in the success path
only; catch-free in Core (see Rule 1/1b above).

**Session Registry:** DI-injected `PresentationSessionRegistry registry` parameter on static MCP
tool methods is correctly EXCLUDED from the MCP JSON schema by SDK 1.3.0 — verified via
`tools/list`. Don't add manual schema suppression for it.

**Two-Layer Shutdown:** `PresentationSessionShutdownService : IHostedService` disposes batches on
host `StopAsync`, AND `Main`'s `finally` calls `registry.DisposeAll()` as an idempotent backstop.
Both layers must stay in sync if session lifecycle changes.

**No CLI/MCP Parity Requirement (Yet):** Unlike Excel, do not assume every MCP action needs a
matching CLI command — the CLI is a placeholder pending a future Generators-based rebuild. Don't
block MCP feature work on CLI parity.

**Generators Deferred:** Do not stand up source generators (mirroring
`ExcelMcp.Generators.Mcp`) until explicitly decided — Core interfaces currently take
`IPresentationBatch batch` directly (no `[ServiceCategory]`/session-id marker attributes the
generator would need). See `.squad/decisions.md` for the full rationale.

**Export Domain:** `ExportSlideToImage`/`ExportAllSlidesToImages` use PowerPoint's native
`Presentation.Export`/`Slide.Export` COM calls — prefer the single multi-slide `Export` call over
a per-slide loop when exporting everything.

## How Path-Specific Instructions Work

GitHub Copilot auto-loads instructions based on files you're editing:

- `tests/**/*.cs` → [Testing Strategy](instructions/testing-strategy.instructions.md)
- `src/PowerPointMcp.Core/**/*.cs`, `src/PowerPointMcp.ComInterop/**/*.cs` → [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- `src/PowerPointMcp.McpServer/**/*.cs` → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- `**` (all files) → [Critical Rules](instructions/critical-rules.instructions.md)

## Agent Skills (`skills/`)

| Skill | File | Target | Best For |
|-------|------|--------|----------|
| **powerpoint-mcp** | `skills/powerpoint-mcp/SKILL.md` | MCP Server | Conversational AI (rich tool schemas) |

**Guidance architecture (single source of truth):**
- `skills/shared/*.md` → manually copied into `skills/powerpoint-mcp/references/` (no build-time
  sync step exists yet — add one before this drifts; see `skills/README.md`).
- NEVER create separate prompt content for guidance that belongs in `skills/shared/`.

## Pre-Commit Hook

Pre-commit runs `scripts/pre-commit.ps1`, which blocks commits if any check fails:

| # | Check | What It Validates |
|---|-------|--------------------|
| 1 | Branch | Never commit to `main` directly |
| 2 | Success Flag | Rule 1: never `Success=true` with `ErrorMessage` |
| 3 | Build | `dotnet build Sbroenne.PowerPointMcp.slnx` succeeds, 0 warnings/errors |
| 4 | Core Tests | Surgical `Feature=`-filtered real-COM integration tests for touched domains |
| 5 | MCP Protocol Tests | `dotnet test tests\PowerPointMcp.McpServer.Tests` (in-memory transport, no COM needed for most cases) |
| 6 | TODO/FIXME Scan | No unresolved `TODO`/`FIXME`/`HACK` markers in changed files |

**Install hook:**
```powershell
# From repo root
Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit
```

## No Confidential Information in Commits/PRs

Never include confidential project names, customer names, or internal references in commit
messages, PR descriptions, or issue text. Use generic descriptions ("a PowerPoint deck", "a chart
shape") instead.
