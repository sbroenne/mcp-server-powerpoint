---
layout: default
title: "Contributing"
description: "How to contribute to PowerPoint MCP Server development. Guidelines for pull requests, code style, and community participation."
permalink: /contributing/
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <h1 class="hero-title">Contributing</h1>
      <p class="hero-subtitle">How to contribute to PowerPoint MCP Server development</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">

## Getting Started

PowerPoint MCP Server is a Windows-only, .NET 10 project that drives a live PowerPoint desktop
instance over COM (`Microsoft.Office.Interop.PowerPoint`). Contributions require a Windows
machine with Microsoft PowerPoint installed for building and running the real-COM integration
tests.

```powershell
git clone https://github.com/sbroenne/mcp-server-powerpoint.git
cd mcp-server-powerpoint
dotnet build
```

## Architecture

The codebase is layered: `ComInterop` → `Core` → `CLI` / `MCP Server`, following the same
architecture as its sibling project, [mcp-server-excel](https://github.com/sbroenne/mcp-server-excel).
See `.github/copilot-instructions.md` in the repository for the full architectural conventions,
including the `Success`/`ErrorMessage` result invariant and the real-COM integration-test
philosophy.

## Reporting Issues

Please use [GitHub Issues](https://github.com/sbroenne/mcp-server-powerpoint/issues) for bug
reports and feature requests. Include:

- PowerPoint version and Windows version
- Steps to reproduce
- Expected vs. actual behavior
- Relevant logs (stderr output from the MCP server, if applicable)

## Pull Requests

- Keep changes focused and scoped to a single concern
- Follow the existing code style and layering conventions
- New behavior needs a corresponding test where practical (real-COM integration tests for
  anything that touches PowerPoint)
- Update documentation (including this site, under `gh-pages/`) when behavior changes

## Code of Conduct

Be respectful and constructive. This is a small open-source project maintained in spare time —
patience with review turnaround is appreciated.

</div>
