# MCP Server for PowerPoint

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

31 tools across 10 domains:

| Domain | Tools |
| --- | --- |
| Presentation | create, open, save, close, list sessions |
| Slide | add, count, delete |
| Shape | add rectangle / text box, count, delete, position, size |
| Text | set/get text, font size, bold, font color |
| Table | add table, set/get cell text |
| Notes | set/get speaker notes |
| Layout | set/get slide layout |
| Image | add picture |
| Chart | add chart, get chart data |
| Export | export a slide / all slides to images (visual verification) |

## How it works

`ComInterop` (STA thread + OLE message filter) → `Core` (domain command
classes) → `McpServer` (stdio host with an in-process session registry). Each
open presentation is a session identified by a `sessionId`; tools operate on a
session until it is closed.

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
