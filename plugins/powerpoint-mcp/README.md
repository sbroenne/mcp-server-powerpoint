# PowerPoint MCP Agent Plugin

Windows-only PowerPoint automation through the PowerPointMcp Model Context Protocol server.

The plugin installs the `powerpoint-mcp` Agent Skill and launches the public
`@sbroenne/mcp-server-powerpoint` npm package through `npx`.

## Prerequisites

- Windows 10 or later
- Microsoft PowerPoint desktop
- Node.js 18 or later with npm/npx
- Network access for npm installation and updates

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

## Runtime launch

The plugin runs `npx -y @sbroenne/mcp-server-powerpoint@latest`. npm selects the native
Windows x64 or ARM64 runtime that matches Node.js and manages its normal package cache.

The plugin contains no PowerPoint documents and does not upload presentation content. The MCP
server drives the locally installed PowerPoint desktop application through Microsoft COM.
