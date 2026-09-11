---
applyTo: ".github/workflows/**,.github/github-app.yml,**/*.csproj,global.json,Directory.Build.*,Directory.Packages.props,scripts/**/*.ps1"
excludeAgent: "code-review"
---

# Build and release constraints

- Keep SDK setup compatible with `global.json`; preserve warnings-as-errors and
  analyzers. Use existing CI and pre-commit checks, not substitute scripts.
- `ci.yml` runs without PowerPoint. Real-COM tests stay local on Windows with
  desktop PowerPoint. Do not remove that coverage or pretend a build covers COM.
- Preserve pre-commit selection of docs-only, tooling, and runtime changes.
  MCP tests marked `RequiresPowerPoint=true` need PowerPoint even though most
  MCP protocol tests do not.
- Local builds call `scripts/Stop-PowerPointMcpProcesses.ps1`. Preserve tracked
  PID-plus-start-time ownership. In shared environments, use
  `-p:PowerPointMcpSkipCleanup=true` to avoid build-triggered cleanup; this is not
  permission to bypass hooks or disable test cleanup.
- `release.yml` owns versions, changelog generation, and publication. Never
  dispatch it as a test.
- Skill guidance belongs in `skills/shared/`. `scripts/Build-AgentSkills.ps1`
  synchronizes references into both skill packages; follow `skills/README.md`.
  Do not edit copies alone or assume an ordinary solution build packages skills.

Procedures: [release strategy](../../docs/RELEASE-STRATEGY.md),
[plugin publication](../workflows/docs/publish-plugins-setup.md), and
[agent setup](../../docs/AGENT-CONFIGURATION.md).
Hook setup and check selection: [development guide](../../docs/DEVELOPMENT-GUIDE.md).
