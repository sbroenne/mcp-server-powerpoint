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

The project ships **two equal, first-class entry points** — an MCP Server and a CLI — that share
one `Core` codebase and are kept in parity by source generators, mirroring `mcp-server-excel`'s
Unified Service Architecture.

**Core Layers:**
1. **ComInterop** (`src/PowerPointMcp.ComInterop`) - STA thread + OLE message filter + channel-based
   work queue (`PresentationBatch`), ported from `mcp-server-excel`'s `ExcelBatch` pattern.
2. **Core** (`src/PowerPointMcp.Core`) - PowerPoint domain commands, one folder per domain
   (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Master, Animation, Image, Chart,
   Export). Domains other than Presentation carry a `[ServiceCategory]` marker attribute that the
   generators discover.
3. **Generators** (`src/PowerPointMcp.Generators.Mcp`, `src/PowerPointMcp.Generators.Cli`,
   `src/PowerPointMcp.Generators.Shared`) - Roslyn source generators that read `[ServiceCategory]`
   Core interfaces and emit one **action-dispatch tool per domain** for the MCP surface (e.g.
   `slide`, `shape`, `chart`, each taking an `operation` parameter like `"chart.add-chart"`) and
   one `pptcli {category} {action}` command per operation for the CLI. `Presentation` (session
   lifecycle) and its `ApplyTemplate`/`GetThemeName` methods stay hand-written
   (`PresentationTools.cs`) since they don't fit the per-session action-dispatch shape.
4. **Service** (`src/PowerPointMcp.Service`) - `PowerPointMcpService`: the shared session registry
   + dispatch layer both entry points call into. **McpServer** hosts it **in-process** (no pipe,
   via `ServiceBridge.ForwardToService`); **CLI** (`pptcli`) talks to it via a **separate
   background daemon process** over a named pipe (`ServiceClient`/`IPowerPointDaemonRpc`,
   auto-started on first `session open`/`session create`), so sessions persist across CLI
   invocations without paying PowerPoint's ~90-150s launch cost every command.
5. **McpServer** (`src/PowerPointMcp.McpServer`) - Model Context Protocol stdio host. 18 tools
   total: 7 hand-written session-lifecycle/template tools
   (`create_presentation`/`open_presentation`/`save_presentation`/`close_presentation`/
   `list_sessions`/`apply_template`/`get_theme_name`) + 11 generated action-dispatch tools (one per
   remaining Core domain), covering ~98 operations. See
   `tests/PowerPointMcp.McpServer.Tests/Integration/McpProtocolTests.cs`'s `ExpectedToolNames` for
   the ground-truth tool list.
6. **CLI** (`src/PowerPointMcp.CLI`) - `pptcli`, built on the same generators as the MCP surface
   (`CliCommandRegistration.RegisterCommands`) plus hand-written `session`/`service` command
   branches for daemon lifecycle. Full parity with the MCP surface (Export is exposed here too).

MCP Server and CLI run as **separate processes**, each managing its own PowerPoint instance — they
do **not** share live sessions with each other, only the same `Core`/`Service` codebase.

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

## PIA-First COM Access (CRITICAL)

All PowerPoint COM operations must use strongly typed members from the embedded
`Microsoft.Office.Interop.PowerPoint` PIA whenever the restored PIA exposes the required API.
Use typed Office/PowerPoint enums instead of raw integer constants.

Only use `dynamic`, reflection, raw dispatch, or numeric COM enum values after confirming from the
restored interop metadata that no typed PIA member/type exists. Keep such exceptions narrowly
scoped, document the missing PIA surface in code, and cover the behavior with a real-COM
integration test. Convenience is never a reason to bypass the PIA.

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

**Session Registry / Service DI:** The DI-injected `PresentationSessionRegistry registry` parameter
(hand-written session-lifecycle tools) and `PowerPointMcpService service` parameter (generated
action-dispatch tools) are both correctly EXCLUDED from the MCP JSON schema by SDK 1.3.0 —
verified via `tools/list`. Don't add manual schema suppression for either.

**Two-Layer Shutdown:** `PresentationSessionShutdownService : IHostedService` disposes batches on
host `StopAsync`, AND `Main`'s `finally` calls `registry.DisposeAll()` as an idempotent backstop.
Both layers must stay in sync if session lifecycle changes.

**No CLI/MCP Parity Gaps:** Both entry points are built from the same generators against the same
`[ServiceCategory]` Core interfaces, so a new Core operation is automatically available on both
surfaces once its domain's generator runs — there is no placeholder/parity-gap state to track
anymore. Export is exposed identically on both.

**Generators Are Live:** `PowerPointMcp.Generators.Mcp`/`PowerPointMcp.Generators.Cli` (mirroring
`ExcelMcp.Generators.Mcp`/`.Cli`) read `[ServiceCategory]`-attributed Core interfaces and emit the
action-dispatch MCP tools and `pptcli` commands. Adding a new domain means adding the
`[ServiceCategory]` attribute + interface/implementation to Core — the generators pick it up
automatically; do not hand-write a new tool class for it.

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
