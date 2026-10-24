# PowerPoint MCP Server — Claude Desktop Bundle

Let Claude Desktop create, edit, and export Microsoft PowerPoint presentations.
Build slides, format text and shapes, add charts and tables, insert images, and
write speaker notes. Export PDFs for sharing or slide images so Claude can
review its work.

This [PowerPoint MCP Server](https://powerpointmcpserver.dev) bundle connects
Claude Desktop to the installed PowerPoint desktop application on Windows.

> **Windows only.** Requires Microsoft PowerPoint (desktop) and Node.js 18 or later with npm/npx.

## Install

1. Download `powerpoint-mcp-<version>.mcpb` from the
   [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases).
2. Double-click the file, or drag-and-drop it onto the Claude Desktop window.
3. Confirm the installation prompt.

Claude can now create and edit PowerPoint decks directly.

## What's inside

A metadata-only configuration that runs `npx -y @sbroenne/mcp-server-powerpoint@latest`,
plus the manifest, license, and changelog. npm selects the native Windows x64 or ARM64 runtime.
The server exposes 32
tools (215 operations across 17 domains) — see the
[documentation](https://powerpointmcpserver.dev) for the full list.
Linked pictures are managed through the generated `shape` actions `get-link-info`, `update-link`,
`break-link`, and `set-link-auto-update`.
Connectors can be free-floating or attached to shape connection sites with `add-connector` and
`add-attached-connector`.
Shapes can be combined with PowerPoint's boolean merge operations through the `shape` tool's
`merge` action.

## Building locally

```pwsh
cd mcpb
./Build-McpBundle.ps1
```

Output is written to `mcpb/artifacts/powerpoint-mcp-<version>.mcpb`.
