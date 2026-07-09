# Changelog

## [0.0.2] - 2026-07-09

### Patch Changes

- [#7](https://github.com/sbroenne/mcp-server-powerpoint/pull/7) [`4bd5377`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4bd5377c359017f75a7af74c0036aa76e2396a53) Thanks [@sbroenne](https://github.com/sbroenne)! - Replaced manual `CHANGELOG.md` editing and the fragile awk/sed-based release-notes
  extraction with [changesets](https://github.com/changesets/changesets): contributors
  now add a small `.changeset/*.md` fragment describing their change, CI enforces one
  is present (or the `skip-changelog` label), and `scripts/Build-Changelog.ps1`
  compiles pending fragments into `CHANGELOG.md` and the GitHub Release body at
  release time. See `docs/RELEASE-STRATEGY.md` for the full process.

- [#7](https://github.com/sbroenne/mcp-server-powerpoint/pull/7) [`4bd5377`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4bd5377c359017f75a7af74c0036aa76e2396a53) Thanks [@sbroenne](https://github.com/sbroenne)! - Added PowerPoint-branded icon assets for the VS Code extension, the MCPB bundle,
  and the gh-pages documentation site, replacing the generic/missing icons used
  previously.

All notable changes to the PowerPoint MCP Server are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.1] - 2026-07-09

### Added

- **MCP server** with 31 tools across 10 domains: Presentation (create, open,
  save, close, list sessions), Slide (add, count, delete), Shape (rectangle, text
  box, count, delete, position, size), Text (set/get text, font size, bold, font
  color), Table (add, set/get cell text), Notes (set/get), Layout (set/get),
  Image (add picture), Chart (add, get data), and Export (slide / all slides to
  image) for visual verification.
- **Live PowerPoint COM automation** via an STA thread with an OLE message
  filter; one long-lived session per open presentation, addressed by `sessionId`.
- **Resilient shutdown**: sessions are disposed on host shutdown with exponential
  backoff so no `POWERPNT.exe` process is orphaned.
- **Non-blocking `create_presentation`**: creates and keeps the deck open,
  returning a `sessionId` immediately (~2s instead of ~90-210s).
- **Distribution**: NuGet .NET tools (`Sbroenne.PowerPointMcp.McpServer` →
  `mcp-powerpoint`, `Sbroenne.PowerPointMcp.CLI` → `pptcli`), MCP Registry
  manifest, and a Claude Desktop MCPB bundle.
- **Agent skill pack** and documentation site at
  [powerpointmcpserver.dev](https://powerpointmcpserver.dev).

[0.0.1]: https://github.com/sbroenne/mcp-server-powerpoint/releases/tag/v0.0.1
