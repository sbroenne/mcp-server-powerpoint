---
title: Installation Guide
description: Complete installation instructions for PowerPoint MCP Server — NuGet .NET tools, standalone executable, VS Code extension, MCPB bundle, MCP Registry and Agent Skills.
keywords: "install PowerPoint MCP Server, PowerPoint MCP setup Windows, dotnet tool install PowerPoint, mcp-powerpoint install, pptcli install, PowerPoint MCP VS Code extension"
# The `howto:` list below is rendered as HowTo JSON-LD by overrides/main.html.
# Each `anchor` must match a heading on this page.
howto_name: How to install PowerPoint MCP Server on Windows
howto_time: PT10M
howto:
  - name: Check the prerequisites
    anchor: prerequisites
    text: Confirm you are on Windows with Microsoft PowerPoint 2016 or later (desktop) installed and activated. The .NET 10 runtime is needed only for the dotnet tool install path.
  - name: Choose an install path
    anchor: choose-your-install-path
    text: Install the MCP Server or CLI as a global .NET tool, download the self-contained standalone executable, install the VS Code extension, or open the MCPB bundle with Claude Desktop.
  - name: Point your MCP client at the server
    anchor: manual-configuration
    text: Run npx add-mcp "mcp-powerpoint" --name powerpoint-mcp, or add the mcp-powerpoint command to your client's MCP settings manually.
  - name: Install the agent skill
    anchor: agent-skills
    text: Run npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp to give your AI assistant workflow guidance beyond the raw tool schemas.
  - name: Verify the install
    anchor: verifying-your-install
    text: Ask your AI assistant to create a presentation and export a slide as an image. A rendered PNG of a real PowerPoint slide confirms the setup works.
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

## Client-specific setup

Step-by-step guides with copy-paste config blocks:

- [Claude Desktop](claude-desktop.md) — MCPB bundle or `claude_desktop_config.json`
- [VS Code & GitHub Copilot](vs-code.md) — extension or `mcp.json`
- [Cursor](cursor.md) — `.cursor/mcp.json`

Claude Code, Windsurf, Cline and Continue follow the same pattern — point them
at the `mcp-powerpoint` command.

## Troubleshooting

The most common problems:

- **"Microsoft PowerPoint is not installed"** — the full desktop application
  must be installed and activated. PowerPoint Online and PowerPoint Mobile
  cannot be automated.
- **.NET tool not found on PATH** — restart your terminal *and* your MCP client
  after `dotnet tool install --global`.
- **Lingering `POWERPNT.EXE` processes** — usually normal. Office's own
  post-quit cleanup can take up to ~90–200 seconds.

See the [troubleshooting guide](troubleshooting.md) for the full list, including
COM HRESULT codes and CLI daemon errors.

## More information

- [Complete Feature Reference](features.md) — all 13 tools (141 operations) across 13 domains
- [MCP Server Documentation](mcp-server.md) — MCP tool reference
- [CLI Documentation](cli.md) — CLI command reference
- [Agent Skills](skills.md) — AI guidance for Claude Code, Cursor, Windsurf and more
- [FAQ](faq.md) — requirements, client compatibility and behaviour
