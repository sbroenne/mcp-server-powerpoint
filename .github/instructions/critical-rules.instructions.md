---
applyTo: "**"
---

# CRITICAL RULES - MUST FOLLOW

> **NON-NEGOTIABLE rules for all PowerPointMcp development**

## Rule 1: Success/ErrorMessage Invariant

**NEVER `Success = true` with a non-null `ErrorMessage`.** Set `Success` only in the success path;
the field defaults to `false` (or is set explicitly `false`) alongside `ErrorMessage` on every
failure path.

**Why Critical:** Confuses LLM callers and downstream consumers of the MCP JSON payload — a
non-null error message on a "successful" result is silently ignored by most clients.

## Rule 1b: No Exception Suppression in Core

**NEVER wrap `batch.Execute()` in a try-catch that swallows the exception and returns an error
result.** Let exceptions propagate naturally out of Core commands — `IPresentationBatch.Execute`'s
underlying `TaskCompletionSource` plumbing (in `PresentationBatch`, ComInterop layer) already
surfaces them to the caller.

The **only** place a catch-all belongs is the MCP tool boundary
(`PowerPointToolsBase.ExecuteToolAction` in `src/PowerPointMcp.McpServer/Tools/`), which logs the
HResult to stderr and serializes a structured error so the MCP host never crashes.

```csharp
// CORRECT — Core: exceptions propagate
public ShapeOperationResult AddRectangle(IPresentationBatch batch, int slideIndex, ...)
{
    return batch.Execute((ctx, ct) => {
        // COM access — no try/catch here
    });
}

// WRONG — Core: swallowing and returning a fabricated error result
public ShapeOperationResult AddRectangle(IPresentationBatch batch, int slideIndex, ...)
{
    try { return batch.Execute(...); }
    catch (Exception ex) { return new ShapeOperationResult { Success = false, ErrorMessage = ex.Message }; }
}
```

**Expected, caller-correctable failures ARE validated up front and returned as `Success=false`**
without throwing — e.g. an out-of-range `slideIndex`, a missing file, or an unknown `sessionId`.
The distinction: validation failures are known preconditions checked BEFORE touching COM;
unexpected COM/runtime exceptions are allowed to propagate and are only ever caught at the MCP
tool boundary.

## Rule 30: Real-COM Integration Tests Only, Strict TDD

**NEVER write unit tests with mocked COM objects for Core commands.** Every Core test drives a
real PowerPoint desktop instance via `PresentationSession.BeginBatch`/`CreateNew`. Follow strict
TDD: write a failing test first (red), watch it fail for the right reason, implement the minimal
fix (green), then verify.

**Enforcement:**
- Tests are tagged `[Trait("Category", "Integration")]` + `[Trait("Feature", "<Domain>")]`.
- Tests run serialized: `xunit.runner.json` sets `maxParallelThreads: 1` — concurrent PowerPoint
  process launches are not supported and will cause flaky failures or hangs.
- Use `--filter "Feature=<Domain>"` for surgical, fast feedback instead of running the whole suite
  after every small change.
- MCP transport tests (`tests/PowerPointMcp.McpServer.Tests`) may use the SDK's in-memory pipe
  transport for protocol-only assertions (no COM launch) — only session-lifecycle round-trip
  tests need real COM at the MCP layer.

## Rule: 1-Based Indexing Everywhere

`slideIndex`, `shapeIndex`, and table `row`/`column` are 1-based throughout Core, the MCP tool
surface, and tests — matching PowerPoint's native COM object model. Never introduce 0-based
indexing in new code; a 0 or negative index is an expected validation failure
(`Success=false`), not something to special-case as valid.

## Rule: Session Lifecycle Discipline

- Every session (`PresentationSessionRegistry` entry) MUST be reachable from
  `PresentationSessionRegistry.DisposeAll()` at host shutdown — never create a code path that
  starts a `PresentationBatch` outside the registry's tracking.
- `presentation(action: "close", sessionId: ...)` is asynchronous by design: it removes the
  session from the registry immediately and disposes the batch (and its PowerPoint process) on a
  background task. Do not make it block on process exit — Office's own cleanup can legitimately
  take minutes.

## Rule: No Confidential Information in Commits/PRs/Issues

Never include confidential project names, customer names, or internal file paths that identify a
specific customer engagement in commit messages, PR descriptions, or issue text. Use generic
descriptions ("a PowerPoint deck", "a chart shape") instead.

## Rule: Never Commit Automatically

Never commit, push, or merge without explicit user approval. No background or silent commits.

## Quick Reference (Grouped by Context)

**Every Edit:**
| Rule | Action | Why Critical |
|------|--------|---------------|
| Success flag | NEVER `Success=true` with `ErrorMessage` | Confuses callers, silent failures |
| No exception wrapping | Never catch-and-return-error inside Core | Preserves stack context, avoids double-wrapping |
| 1-based indexing | Never introduce 0-based indices | Matches COM object model, avoids silent bugs |

**When Writing Code:**
| Rule | Action | Why Critical |
|------|--------|---------------|
| TDD | Write the failing test FIRST → red → implement → green | Proves the test actually catches the bug |
| Integration tests only | NEVER mock COM — real PowerPoint only | Unit tests prove nothing for COM interop |
| COM cleanup | Use try/finally for any manually-acquired COM object refs | Prevents leaks |

**Before Commit:**
| Rule | Action |
|------|--------|
| Build | `dotnet build Sbroenne.PowerPointMcp.slnx` — 0 warnings, 0 errors |
| Tests | Run `Feature=`-filtered tests for the domain(s) you touched |
| Pre-commit hook | `scripts/pre-commit.ps1` must pass |
| No confidential info | Scrub commit messages/PR text before submitting |
