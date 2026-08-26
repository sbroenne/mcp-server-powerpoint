# PowerPoint MCP Agent Plugin

Windows-only PowerPoint automation through the PowerPointMcp Model Context Protocol server.

The plugin installs the `powerpoint-mcp` Agent Skill, MCP configuration, and a small bootstrap.
On first use, the bootstrap downloads the matching self-contained runtime from the latest
[`sbroenne/mcp-server-powerpoint`](https://github.com/sbroenne/mcp-server-powerpoint/releases)
release. The runtime is cached per user and verified against the release `SHA256SUMS` file.

## Prerequisites

- Windows 10 or later
- Microsoft PowerPoint desktop
- Windows PowerShell 5.1 or PowerShell 7
- Network access for the first runtime download

After one verified download, the cached runtime remains available during temporary network
outages. A cold cache fails with a clear error instead of running unverified code.

## Installation

```text
copilot plugin install sbroenne/mcp-server-powerpoint-plugins/powerpoint-mcp
```

## What it supports

- Presentation create, open, test, list, and save-on-close
- Slides, sections, comments, backgrounds, and imports
- Shapes, placeholders, text formatting, and hyperlinks
- Tables, native charts, SmartArt, images, and speaker notes
- Layouts, masters, page setup, accessibility, and animation
- PDF and image export for delivery and visual verification

The MCP server exposes one action-dispatch tool per PowerPoint domain. Tool schemas provide exact
action names and parameters; the bundled skill adds lifecycle and workflow guidance.

## Runtime safety

- Release archives and cached archives are checked against `SHA256SUMS`.
- Downloads use a temporary `.part` file and staged extraction.
- A per-user mutex serializes installs.
- An existing runtime is replaced only after the new runtime passes version checks.
- A runtime currently in use is never overwritten.
- Release checks use GitHub authentication when `GH_TOKEN` or `GITHUB_TOKEN` is available.

The plugin contains no PowerPoint documents and does not upload presentation content. The MCP
server drives the locally installed PowerPoint desktop application through Microsoft COM.
