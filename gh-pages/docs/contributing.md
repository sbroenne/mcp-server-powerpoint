---
title: Contributing
description: How to contribute to PowerPoint MCP Server development. Guidelines for pull requests, code style and community participation.
---

# Contributing

## Getting started

PowerPoint MCP Server is a Windows-only, .NET 10 project that drives a live
PowerPoint desktop instance over COM
(`Microsoft.Office.Interop.PowerPoint`). Contributions require a Windows
machine with Microsoft PowerPoint installed for building and running the
real-COM integration tests.

```powershell
git clone https://github.com/sbroenne/mcp-server-powerpoint.git
cd mcp-server-powerpoint
dotnet build
```

### Package sources and dependency updates

Restore commands inherit your local NuGet, npm, and pip settings. The repository
does not clear NuGet sources or force a download registry. Configure any required
package feed and credentials locally; do not commit machine-specific settings.

NuGet versions are maintained in `Directory.Packages.props`. Each npm project
has its own `package.json` and `package-lock.json`; use `npm ci` to install the
locked versions. Each npm project's `.npmrc` sets
`omit-lockfile-registry-resolved=true`, so normal npm installs and updates omit
registry download URLs while retaining versions and integrity checks. This setting
does not change your registry, proxy, or credentials. Nested npm projects need
their own `.npmrc`; add the same setting when creating a new npm project.

CI and the commit hook reject fixed download URLs in tracked npm lockfiles.
The hook checks the staged content, even if your working copy has already been
fixed. To repair a lockfile, run
`npm install --package-lock-only --ignore-scripts` in its project and stage the
regenerated file. Run `pwsh -File scripts/check-npm-lockfiles.ps1` from the
repository root to check all tracked lockfiles locally.

The extension's `@types/vscode` stays on the minor version declared by
`engines.vscode`, so dependency updates do not silently require a newer VS Code.
Update both together when intentionally raising the minimum supported version.
Dependabot checks NuGet, npm (including the intro video), documentation packages,
and GitHub Actions weekly.

## Architecture

The codebase is layered: `ComInterop` → `Core` → `CLI` / `MCP Server`,
following the same architecture as its sibling project,
[mcp-server-excel](https://github.com/sbroenne/mcp-server-excel). See
`.github/copilot-instructions.md` in the repository for the full
architectural conventions, including the `Success`/`ErrorMessage` result
invariant and the real-COM integration-test philosophy.

## Reporting issues

Please use [GitHub Issues](https://github.com/sbroenne/mcp-server-powerpoint/issues)
for bug reports and feature requests. Include:

- PowerPoint version and Windows version
- Steps to reproduce
- Expected vs. actual behavior
- Relevant logs (stderr output from the MCP server, if applicable)

## Pull requests

- Keep changes focused and scoped to a single concern
- Follow the existing code style and layering conventions
- New behavior needs a corresponding test where practical (real-COM
  integration tests for anything that touches PowerPoint)
- Update documentation (including this site, under `gh-pages/`) when
  behavior changes

## Code of conduct

Be respectful and constructive. This is a small open-source project
maintained in spare time — patience with review turnaround is appreciated.
