# PowerPointMcp repository rules

GitHub Copilot is the primary coding agent, locally and on GitHub. Use Windows,
PowerShell, and the SDK selected by `global.json`. Desktop PowerPoint is required
for COM tests; GitHub-hosted runners do not have it.

MCP Server and `pptcli` are equal entry points: behavior, defaults, validation,
results, and documentation must agree. Read
[critical rules](instructions/critical-rules.instructions.md) for every change.
Read only the additional instructions matching the work:

| Work | Instructions |
| --- | --- |
| Runtime, contracts, generators | [Architecture](instructions/architecture-patterns.instructions.md) |
| MCP tools, schemas, MCP generators | [MCP guide](instructions/mcp-server-guide.instructions.md) |
| Tests or behavior changes | [Testing](instructions/testing-strategy.instructions.md) |
| Builds, scripts, workflows, releases | [Development workflow](instructions/development-workflow.instructions.md) |
| Agent instructions | [Instruction maintenance](instructions/meta.instructions.md) |
| Skills | `skills/CLAUDE.md` and `skills/README.md` |

## Implementation

- Core `[ServiceCategory]` interfaces drive generated routing. Change contracts
  and generators, not emitted code. Follow changes through both entry points,
  tests, help, and shared guidance.
- `sbroenne/mcp-server-excel` is the architectural reference. Before architectural
  changes, check its existing solution. Preserve PowerPoint-specific differences
  such as PIA embedding and the absence of `Application.ScreenUpdating`.
  Flag shared bugs for Excel or `mcp-windows` without changing those repos.
- Keep rules in their linked files, not duplicated here. Tool and operation
  counts come from the generated skill manifest, `PresentationToolAction`, and
  `McpProtocolTests.ExpectedToolNames`, not remembered inventories.

## Build and validation

```powershell
dotnet restore Sbroenne.PowerPointMcp.slnx
dotnet build Sbroenne.PowerPointMcp.slnx -c Release --no-restore
```

Build with zero warnings. Run focused checks from the
[testing strategy](instructions/testing-strategy.instructions.md) and existing
gates in `scripts/pre-commit.ps1` / `.github/workflows/ci.yml`; do not invent
replacement audits. `scripts/check-doc-counts.ps1` needs a current Release build
in this worktree. Documentation/configuration-only changes do not need synthetic
tests or PowerPoint launches; validate their links, syntax, and affected behavior.
Report unavailable COM checks as not run, never as passed.

## Git and release

- Never commit to `main`, force-push, or bypass hooks. Report hook blockers.
- Commit/push only when explicitly authorized. A GitHub coding-agent assignment
  explicitly requesting a PR authorizes its delivery commits and PR, not merging
  or publishing. Local edits alone do not authorize commits.
- Verify `gh auth status` uses the personal `sbroenne` account for this namespace.
  An inherited `GH_TOKEN` can override the stored account; never print tokens.
- User-visible changes need a changeset. Internal/docs/tests/CI changes use the
  `skip-changelog` PR label. Versions and `CHANGELOG.md` are release-generated.
- Plugin publication follows
  [the publication guide](workflows/docs/publish-plugins-setup.md);
  the published repository is output-only.

Agent setup and local/GitHub limitations:
[docs/AGENT-CONFIGURATION.md](../docs/AGENT-CONFIGURATION.md).
Background, examples, hook setup, and the guidance migration map:
[development guide](../docs/DEVELOPMENT-GUIDE.md).
