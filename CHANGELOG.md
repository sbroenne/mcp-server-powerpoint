# Changelog

All notable changes to the PowerPoint MCP Server are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/sbroenne/mcp-server-powerpoint/commits/main
