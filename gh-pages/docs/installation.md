---
title: Installation Guide
description: Complete installation instructions for PowerPoint MCP Server — NuGet .NET tools, standalone executable, VS Code extension, MCPB bundle, MCP Registry and Agent Skills.
---

# Installation Guide

## Prerequisites

- **Windows** (this project drives a real, local PowerPoint desktop process
  via COM — there is no cross-platform or headless mode)
- **Microsoft PowerPoint** (desktop) installed and licensed
- **.NET 10 runtime** — only required for the NuGet `dotnet tool` install
  path; the standalone executable and MCPB bundle are self-contained

## Choose your install path

=== "NuGet .NET Tool — MCP Server"

    Install the MCP server as a global .NET tool. Requires the .NET 10
    runtime.

    ```powershell
    dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
    ```

    Runs as the `mcp-powerpoint` command once installed.

=== "NuGet .NET Tool — CLI"

    Install the token-efficient CLI for coding agents. Requires the .NET 10
    runtime.

    ```powershell
    dotnet tool install --global Sbroenne.PowerPointMcp.CLI
    ```

    Runs as the `pptcli` command once installed.

=== "Standalone executable"

    Download a self-contained build from GitHub Releases — no .NET runtime
    install needed.

    ```powershell
    # https://github.com/sbroenne/mcp-server-powerpoint/releases/latest
    # Extract and run mcp-powerpoint.exe / pptcli.exe directly
    ```

=== "VS Code Extension"

    One-click install that auto-configures the MCP server for GitHub Copilot
    Chat in VS Code.

    [Install Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp){ .md-button .md-button--primary }

=== "MCPB Bundle — Claude Desktop"

    One-click install for Claude Desktop via the MCP Bundle (`.mcpb`) format.

    ```powershell
    # https://github.com/sbroenne/mcp-server-powerpoint/releases/latest
    # Download the .mcpb file and open it with Claude Desktop
    ```

=== "MCP Registry"

    Discoverable through the official MCP Registry for any MCP-compatible
    client that supports registry-based discovery.

## Manual configuration

After installing via NuGet or the standalone executable, point your MCP
client at the server:

```powershell
# After installing either way, auto-configure supported coding agents
npx add-mcp "mcp-powerpoint" --name powerpoint-mcp
```

Or configure manually in your client's MCP settings (VS Code `mcp.json`,
Claude Desktop config, etc.) to run `mcp-powerpoint` (NuGet tool install) or
the path to the extracted `mcp-powerpoint.exe` (standalone download).

## Agent Skills

Skills give AI assistants workflow guidance beyond raw tool schemas —
strongly recommended, especially for the CLI, and useful even for the MCP
server's richer tool discovery:

```powershell
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

!!! tip
    The VS Code extension installs this skill automatically to
    `~/.copilot/skills/powerpoint-mcp/`. Manual `npx skills add` is for other
    clients (Claude Code, Cursor, Windsurf, etc.).

## Verifying your install

Ask your AI assistant something like:

> "Create a new PowerPoint presentation with a title slide, then export it
> as an image so I can see it."

If the assistant can call `presentation(action="create", filePath=...)`,
`slide(action="add-blank", session_id=...)`,
`textframe(action="set-text", session_id=..., slide_index=..., shape_index=..., text=...)`, and
`export(action="export-slide-to-image", session_id=..., slide_index=..., output_path=...)` and
you get back a rendered PNG of a real PowerPoint slide, you're set up correctly.

## Troubleshooting

- **"PowerPoint is not installed" errors** — this project requires the full
  PowerPoint desktop application (not PowerPoint Online / Mobile) to be
  installed and activated on the same Windows machine running the MCP
  server.
- **Lingering `POWERPNT.EXE` processes** — sessions are cleaned up on
  `presentation(action="close", sessionId=...)` and on MCP server shutdown; if a process lingers
  after a crash, close it from Task Manager.
- **.NET tool not found on PATH** — restart your terminal after
  `dotnet tool install --global` so the updated `PATH` is picked up.

## More information

- [Complete Feature Reference](features.md) — all 13 tools (134 operations) across 13 domains
- [MCP Server Documentation](mcp-server.md) — MCP tool reference
- [CLI Documentation](cli.md) — CLI command reference
- [Agent Skills](skills.md) — AI guidance for Claude Code, Cursor, Windsurf and more
