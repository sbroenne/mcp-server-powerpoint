# MCP Server for PowerPoint

<!-- mcp-name: io.github.sbroenne/mcp-server-powerpoint -->
mcp-name: io.github.sbroenne/mcp-server-powerpoint

A [Model Context Protocol](https://modelcontextprotocol.io) server that automates
Microsoft PowerPoint for AI assistants. It drives a **live PowerPoint desktop
instance** via COM automation, so decks are created and edited exactly as they
would be by hand — then exported to images so an AI can *visually verify* its own
work.

> **Windows only.** Requires Microsoft PowerPoint (desktop) to be installed.

## Install

```pwsh
dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
```

This installs the `mcp-powerpoint` command. Point your MCP client (Claude
Desktop, VS Code, GitHub Copilot, etc.) at it over stdio:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "mcp-powerpoint"
    }
  }
}
```

## Capabilities

**13 tools with 132 operations across 13 domains:**

| Tool | Ops | Coverage |
| --- | --- | --- |
| `presentation` | 12 | create, open, save, close, list, apply-template, get-theme-name, built-in/custom document properties |
| `slide` | 14 | slide lifecycle, solid + gradient backgrounds, sections |
| `shape` | 36 | shape creation, geometry, styling, effects, grouping, naming, hyperlinks |
| `textframe` | 17 | text content, font formatting, alignment, bullets, auto-size |
| `table` | 12 | tables, cell text, row/column edits, cell fill/border, merge |
| `notes` | 2 | set/get speaker notes |
| `layout` | 2 | set/get slide layout |
| `master` | 8 | title/body placeholder fonts, solid + gradient master backgrounds |
| `animation` | 5 | shape effects, transition read/write |
| `image` | 5 | add picture, brightness/contrast, recolor |
| `chart` | 10 | charts, series, chart/axis titles, legend visibility, data replacement |
| `smartart` | 7 | SmartArt insertion and node editing |
| `export` | 2 | export a slide or every slide to images |

Every domain is exposed as a single **action-dispatch tool** taking an `action`
parameter — including `presentation`. Example MCP calls:

- `presentation(action="create", filePath="C:\\Decks\\demo.pptx")`
- `presentation(action="apply-template", sessionId="...", templatePath="C:\\Templates\\brand.potx")`
- `chart(action="add-chart", session_id="...", slide_index=2, ...)`

## How it works

`ComInterop` (STA thread + OLE message filter) → `Core` (domain command classes) →
`PowerPointMcp.Service` (shared service layer) → `McpServer` (stdio host, calling the service
in-process) and `CLI` (talks to the same service via a background named-pipe daemon). Each open
presentation is a session identified by a `sessionId`; tools operate on a session until it is
closed.

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
