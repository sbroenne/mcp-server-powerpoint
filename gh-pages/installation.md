---
layout: default
title: "Installation Guide"
description: "Complete installation instructions for PowerPoint MCP Server - NuGet .NET tools, standalone executable, VS Code extension, MCPB bundle, MCP Registry, and Agent Skills."
permalink: /installation/
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <h1 class="hero-title">Installation Guide</h1>
      <p class="hero-subtitle">Complete installation instructions for PowerPoint MCP Server</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">

## Prerequisites

- **Windows** (this project drives a real, local PowerPoint desktop process via COM — there is
  no cross-platform or headless mode)
- **Microsoft PowerPoint** (desktop) installed and licensed
- **.NET 10 runtime** — only required for the NuGet `dotnet tool` install path; the standalone
  executable and MCPB bundle are self-contained

## Choose Your Install Path

<div class="install-options">

<div class="install-option">
<h3>NuGet .NET Tool <span class="badge">MCP Server</span></h3>
<p>Install the MCP server as a global .NET tool. Requires the .NET 10 runtime.</p>

```powershell
dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
```

Runs as the `mcp-powerpoint` command once installed.
</div>

<div class="install-option">
<h3>NuGet .NET Tool <span class="badge">CLI</span></h3>
<p>Install the token-efficient CLI for coding agents. Requires the .NET 10 runtime.</p>

```powershell
dotnet tool install --global Sbroenne.PowerPointMcp.CLI
```

Runs as the `pptcli` command once installed.
</div>

<div class="install-option">
<h3>Standalone Executable <span class="badge">No .NET required</span></h3>
<p>Download a self-contained build from GitHub Releases — no .NET runtime install needed.</p>

```powershell
# https://github.com/sbroenne/mcp-server-powerpoint/releases/latest
# Extract and run mcp-powerpoint.exe / pptcli.exe directly
```
</div>

<div class="install-option">
<h3>VS Code Extension <span class="badge">Marketplace</span></h3>
<p>One-click install that auto-configures the MCP server for GitHub Copilot Chat in VS Code.</p>

<a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp" class="button-link">Install Extension</a>
</div>

<div class="install-option">
<h3>MCPB Bundle <span class="badge">Claude Desktop</span></h3>
<p>One-click install for Claude Desktop via the MCP Bundle (`.mcpb`) format.</p>

```powershell
# https://github.com/sbroenne/mcp-server-powerpoint/releases/latest
# Download the .mcpb file and open it with Claude Desktop
```
</div>

<div class="install-option">
<h3>MCP Registry</h3>
<p>Discoverable through the official MCP Registry for any MCP-compatible client that supports
registry-based discovery.</p>
</div>

</div>

## Manual Configuration

After installing via NuGet or the standalone executable, point your MCP client at the server:

```powershell
# After installing either way, auto-configure supported coding agents
npx add-mcp "mcp-powerpoint" --name powerpoint-mcp
```

Or configure manually in your client's MCP settings (VS Code `mcp.json`, Claude Desktop config,
etc.) to run `mcp-powerpoint` (NuGet tool install) or the path to the extracted
`mcp-powerpoint.exe` (standalone download).

## Agent Skills

Skills give AI assistants workflow guidance beyond raw tool schemas — strongly recommended,
especially for the CLI, and useful even for the MCP server's richer tool discovery:

```bash
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

> 💡 The VS Code extension installs this skill automatically to `~/.copilot/skills/powerpoint-mcp/`.
> Manual `npx skills add` is for other clients (Claude Code, Cursor, Windsurf, etc.).

## Verifying Your Install

Ask your AI assistant something like:

> "Create a new PowerPoint presentation with a title slide, then export it as an image so I can see it."

If the assistant can call `create_presentation`, `add_slide`, `set_text`, and
`export_slide_to_image` and you get back a rendered PNG of a real PowerPoint slide, you're set up
correctly.

## Troubleshooting

- **"PowerPoint is not installed" errors** — this project requires the full PowerPoint desktop
  application (not PowerPoint Online / Mobile) to be installed and activated on the same Windows
  machine running the MCP server.
- **Lingering `POWERPNT.EXE` processes** — sessions are cleaned up on `close_presentation` and on
  MCP server shutdown; if a process lingers after a crash, close it from Task Manager.
- **.NET tool not found on PATH** — restart your terminal after `dotnet tool install --global` so
  the updated `PATH` is picked up.

## More Information

📖 **[Complete Feature Reference](/features/)** — All 31 tools across 10 domains

📖 **[MCP Server Documentation](/mcp-server/)** — MCP tool reference

📖 **[CLI Documentation](/cli/)** — CLI command reference

🤖 **[Agent Skills](/skills/)** — AI guidance for Claude Code, Cursor, Windsurf, and more

</div>
