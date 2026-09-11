# Development guide

This is the explanatory companion to the short
[repository instructions](../.github/copilot-instructions.md). It preserves
background, examples, and procedures moved out of the always-loaded agent
guidance. Mandatory rules remain in the linked instruction files.

## System map and sister projects

PowerPointMcp automates a live desktop PowerPoint instance on Windows through
the official PowerPoint Primary Interop Assembly (PIA). It serves both coding
agents and automation scripts through two equal entry points: MCP and `pptcli`.

| Area | Responsibility |
| --- | --- |
| `src/PowerPointMcp.ComInterop` | Dedicated STA thread, OLE message filter, and channel-based work queue in `PresentationBatch` |
| `src/PowerPointMcp.Core` | Domain command interfaces, implementations, and result objects |
| `src/PowerPointMcp.Generators` | Shared service dispatch and skill manifest generation from Core contracts |
| `src/PowerPointMcp.Generators.Mcp` | MCP action-dispatch tools |
| `src/PowerPointMcp.Generators.Cli` | CLI settings and command registration |
| `src/PowerPointMcp.Generators.Shared` | Metadata and helpers shared by generators |
| `src/PowerPointMcp.Service` | Session registry and shared command dispatch |
| `src/PowerPointMcp.McpServer` | JSON-RPC stdio host with the service in-process |
| `src/PowerPointMcp.CLI` | `pptcli`, communicating with a persistent service process over a named pipe |

Excel's repository, `sbroenne/mcp-server-excel`, is the architectural reference:
layering, the unified service design, generators, hand-written versus generated
tools, COM cleanup, support scripts, and CI checks. Its `ExcelBatch` design is
the reference for `PresentationBatch`. Investigate differences rather than
assuming a new PowerPoint-specific design is necessary. `mcp-windows` is the
related Windows automation project; flag shared defects without changing it as
part of a PowerPoint task.

Some differences are intentional. PowerPoint has no
`Application.ScreenUpdating`. `ForceEmbedPowerPointInteropTypes` in
`Directory.Build.targets` embeds the PowerPoint and Office interop types, so
there is no equivalent runtime PIA assembly resolver to copy from Excel.
Embedding types removes an assembly dependency, not the need to release COM
references.

See [architecture patterns](../.github/instructions/architecture-patterns.instructions.md)
for class organization, the command interface example, and resource ownership.

## Following an operation through both entry points

Core domains marked `[ServiceCategory]` supply the contract read by the
generators. The MCP surface is one tool per domain with an `action` argument,
not one hand-written tool per operation. CLI registration is generated through
`CliCommandRegistration.RegisterCommands`. Both entry points reach the same
service dispatch and Core command implementations.

`PresentationTools.cs` is the hand-written `presentation` action-dispatch tool.
Its lifecycle, template, and document-property behavior needs special session
handling. Export is available on both entry points; it is not a CLI parity gap.
Adding a generated operation means changing the Core contract and implementation,
not editing generated output or adding a parallel hand-written MCP tool.

The old instruction inventories mixed different generations of tool counts.
Use these sources rather than retaining a snapshot of those numbers:

- The generated `_SkillManifest.g.cs` lists generated domains and operations.
- `PresentationToolAction` supplies the hand-written presentation actions.
- `McpProtocolTests.ExpectedToolNames` defines the expected MCP tool list.
- `scripts/check-doc-counts.ps1` checks advertised counts after a current Release build.

The MCP SDK resolves `PresentationSessionRegistry registry` and
`PowerPointMcpService service` from dependency injection rather than exposing
them as caller arguments. Protocol tests verify this through `tools/list`;
do not add manual schema suppression for these parameters.
See the [MCP guide](../.github/instructions/mcp-server-guide.instructions.md)
for dispatch, validation, serialization, and schema examples.

## Sessions and why they persist

The MCP host is already long-lived, so it hosts the service in-process via
`ServiceBridge.ForwardToService`. CLI invocations are short-lived, so
`ServiceClient` and `IPowerPointDaemonRpc` communicate with a separate background
service using StreamJsonRpc over a named pipe. This preserves CLI sessions and
avoids paying PowerPoint's substantial startup cost on every command. Earlier
guidance observed roughly 90-150 seconds for startup; treat this as an observation,
not a guaranteed timing contract.

MCP and CLI share code, not live sessions or PowerPoint instances.
The MCP lifecycle is:

```text
presentation(action="create", filePath) -> saved file and open sessionId
presentation(action="open", filePath)   -> open sessionId
domain(action=..., session_id=...)     -> operate on the existing session
presentation(action="test", filePath)  -> validate opening; retain no session
presentation(action="close", sessionId, save=true)
                                      -> save, remove session, dispose in background
```

Reuse the session returned by create instead of opening the same file again.
The close call returns without waiting for the PowerPoint process to exit.
Office may take minutes to finish post-Quit cleanup; that delay alone does not
prove a leak or justify force-killing it.

Host shutdown through Ctrl+C, stdin EOF, or normal exit has two cleanup layers:
`PresentationSessionShutdownService.StopAsync` and the `Main` finally block both
call the registry's idempotent `DisposeAll()`. The
[critical rules](../.github/instructions/critical-rules.instructions.md)
define the required tracking and process-ownership safeguards.

For exporting every slide to an image, the architecture guide retains the
preference for one native `Presentation.Export` call over a per-slide loop.

## Why success and exception handling are separate rules

A result reporting success alongside an error message can cause agents and
other callers to ignore the error. Set success only when the operation really
completed; do not initialize it optimistically and then attach an error.

Expected invalid input is different from an unexpected COM failure. Validate
known preconditions before the failing COM call and return an error result.
Unknown-session validation at the MCP boundary, for example, follows this shape:

```csharp
if (!registry.TryGet(sessionId, out var batch))
{
    return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
}
```

In Core, return `batch.Execute(...)` directly rather than wrapping it in a
catch that manufactures an error result. The batch's `TaskCompletionSource`
already carries unexpected exceptions back from the STA thread. Swallowing
them in Core loses diagnostic context and introduces a second error boundary.
`PowerPointToolsBase.ExecuteToolAction` is the MCP boundary that logs the HResult
to stderr and returns a structured error. Stdout is reserved for JSON-RPC.

The [critical rules](../.github/instructions/critical-rules.instructions.md)
retain the success invariant, typed PIA requirements, 1-based indices, and
real-COM regression-test requirement. The
[testing strategy](../.github/instructions/testing-strategy.instructions.md)
retains the red/green sequence, test traits and filters, serialization,
fixture pitfalls, shutdown observations, and protocol-versus-COM distinction.

## Build, checks, and hooks

Use the build commands in the root instructions and targeted tests in the
testing strategy. Use precise file edits and code search for source changes;
PowerShell is appropriate for `dotnet`, Git, and repository scripts. Tool names
vary between agents, so the old editor-specific tool names are not requirements.

The [pre-commit script](../scripts/pre-commit.ps1) is the executable reference
for local checks. Its checks include:

| Check | Selection and purpose |
| --- | --- |
| Branch guard | Blocks direct commits to `main` |
| Process cleanup | Stops owned daemon processes and shuts down build servers |
| Success flag audit | Checks the success/error invariant |
| COM leak, dynamic cast, and interface audits | Run for non-docs changes |
| Release build | Runs for non-docs changes with warnings treated as errors |
| Documentation counts | Always runs; creates a manifest with a one-time build if missing |
| Core tests | Selects `Feature=` tests for changed Core domains |
| MCP tests | Full suite for runtime changes, PowerPoint-free subset for tooling, skipped for docs-only changes |
| Release and skill packaging tests | Runs for non-docs changes |
| Unresolved-marker scan | Scans eligible changed files |

Even a docs-only commit can invoke process cleanup and a first build when the
manifest is absent. Do not confuse manual app commands with the full hook.
The [CI Gate](../.github/workflows/ci.yml) is the server-side reference; it cannot
replace local PowerPoint-dependent checks.

### Installing a hook in a normal checkout or worktree

First inspect `git config --get core.hooksPath` and
`git rev-parse --git-path hooks/pre-commit` from the working checkout.
An unset `core.hooksPath` is normal. Respect any existing hook or hook manager;
do not overwrite it. Worktrees commonly use a `.git` file, not a `.git`
directory, and may share hook storage with another checkout.

Git needs an executable hook launcher. A launcher can invoke the checked-in
PowerShell script using this content:

```sh
#!/bin/sh
exec pwsh -NoProfile -File scripts/pre-commit.ps1
```

Install it at the resolved hook path only after reviewing any existing hook,
and keep it executable where the platform requires that. The old
`Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit` shortcut is not reliable
for worktrees or Git's hook interpreter. Do not bypass a failing hook.

### GitHub account selection

For this repository's namespace, verify that `gh auth status` reports the
personal `sbroenne` account as active. If an inherited `GH_TOKEN` selects a
different account, remove that override only in the command's process before
using `gh auth switch --hostname github.com --user sbroenne`, then recheck.
Never print token values or store them in repository files. Switching accounts
does not authorize committing, pushing, merging, or publishing.

## Authoring skills and shared guidance

The [skills README](../skills/README.md) describes both `powerpoint-mcp` and
`powerpoint-cli`, their distribution, and their build procedure. The former
uses rich MCP schemas; the latter provides a compact command surface for agents
and scripts.

`skills/shared/` is the authoring source of truth. `Build-AgentSkills.ps1`
synchronizes shared references into both skill packages and generates the CLI
command reference from live help. Edit the source, not only a copied reference.
Read [skills/CLAUDE.md](../skills/CLAUDE.md) before changing that area.
An ordinary solution build is not a substitute for the skill packaging procedure.
The old statement that no synchronization tooling exists is obsolete.

## Where the former root guidance went

This map records the consolidation so that shortening the agent prompt does not
hide the information from contributors.

| Former section or lesson | Current home |
| --- | --- |
| Critical files and path-specific loading | Root task table and [agent configuration](AGENT-CONFIGURATION.md) |
| Sister projects and PowerPoint overview | System map above and architecture guide |
| Layer responsibilities, generators, CLI daemon | System map and operation/session explanations above |
| Fixed tool, operation, and project counts | Code-derived sources above; stale inventories replaced |
| Session model and create/test/close behavior | Sessions above and architecture guide |
| 1-based indexing and PIA-first access | Critical rules; resource-management details in architecture guide |
| Rule 1/1b, rationale, and examples | Critical rules, explanation above, and MCP guide |
| Rule 30 and test commands | Testing strategy, including focused filters and actual PowerPoint requirements |
| Tool-selection quick reference | Build section above without agent-specific tool names |
| GitHub account rule | Account-selection procedure above and root authorization rules |
| DI/schema, shutdown, parity, live generators, export lessons | Operation/session sections above, MCP guide, and architecture guide |
| Skills and shared guidance | Skills section above and skills README; obsolete manual-only sync claim corrected |
| Pre-commit table and installation | Hook section above and executable script |
| Confidentiality and commit approval | Critical rules and root Git/release section |

Repeated summaries were consolidated, not treated as separate rules. Incorrect
claims were corrected rather than archived as active instructions: generated
tool counts, PIA embedding managing COM lifetimes, missing skill synchronization,
and the worktree-incompatible hook installation shortcut.
