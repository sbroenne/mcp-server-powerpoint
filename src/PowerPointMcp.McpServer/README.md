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

18 tools with ~98 operations across 12 domains:

| Domain | Tools |
| --- | --- |
| Presentation | create, open, save, close, list sessions |
| Template | apply template, get theme name |
| Slide | add, count, delete, duplicate, move, sections, background color |
| Shape | add rectangle / text box / auto shape / line / connector, position, size, fill, line, rotation, flip, z-order, shadow, group, name, alt text |
| TextFrame | set/get text, font size, bold, color, italic, underline, font name, alignment, bullets |
| Table | add table, set/get cell text, insert/delete rows and columns, cell fill, cell border, merge cells |
| Notes | set/get speaker notes |
| Layout | set/get slide layout |
| Master | title/body placeholder font, background color |
| Animation | add/delete shape effects, get transition, set transition |
| Image | add picture |
| Chart | add chart, get chart data, add series, chart/axis title, legend visibility |
| Export | export a slide / all slides to images (visual verification) |

Most domains are exposed as a single **action-dispatch tool** (e.g. `shape`, `table`) taking an
`operation` parameter, keeping the tool list small for AI assistants while still exposing every
operation above. `Presentation` and `Template` remain individual, hand-written tools.

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
