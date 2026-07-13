# PowerPoint MCP Server — Claude Desktop Bundle

This is the [PowerPoint MCP Server](https://powerpointmcpserver.dev) packaged as
an **MCPB bundle** for one-click installation in **Claude Desktop**.

> **Windows only.** Requires Microsoft PowerPoint (desktop) to be installed.

## Install

1. Download `powerpoint-mcp-<version>.mcpb` from the
   [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases).
2. Double-click the file, or drag-and-drop it onto the Claude Desktop window.
3. Confirm the installation prompt.

Claude can now create and edit PowerPoint decks directly.

## What's inside

A self-contained Windows x64 build of the MCP server (no .NET runtime install
required) plus the manifest, license, and changelog. The server exposes 13
tools (132 operations across 13 domains) — see the
[documentation](https://powerpointmcpserver.dev) for the full list.

## Building locally

```pwsh
cd mcpb
./Build-McpBundle.ps1
```

Output is written to `mcpb/artifacts/powerpoint-mcp-<version>.mcpb`.
