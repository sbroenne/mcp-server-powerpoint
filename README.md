# PowerPointMcp — MCP Server &amp; CLI for Microsoft PowerPoint

[![VS Code Marketplace](https://vsmarketplacebadges.dev/installs-short/sbroenne.powerpoint-mcp.svg?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-powerpoint/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-powerpoint/releases)

[![Build MCP Server](https://github.com/sbroenne/mcp-server-powerpoint/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-powerpoint/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-powerpoint/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-powerpoint/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-powerpoint)](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-powerpoint)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate PowerPoint with AI** — a [Model Context Protocol](https://modelcontextprotocol.io) server and CLI
for live, real-time PowerPoint automation through conversational AI. Sibling project to
[mcp-server-excel](https://github.com/sbroenne/mcp-server-excel), following the same layered
architecture: `ComInterop` → `Core` → `CLI` / `MCP Server`.

**MCP Server for PowerPoint** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to build and
edit real `.pptx` presentations through natural language — slides, shapes, text, tables, charts,
speaker notes, and layouts — no VBA or PowerPoint object-model knowledge required.

**🛡️ Live COM automation, not file parsing** — Most PowerPoint MCP servers manipulate `.pptx`
files offline with libraries like `python-pptx`, or use agent-run scripts with LibreOffice-rendered
thumbnails. This project instead drives a **live, real PowerPoint desktop instance** via
`Microsoft.Office.Interop.PowerPoint` — the official Primary Interop Assembly. That means
true-fidelity rendering, compatibility with an already-open deck, and zero risk of producing a
`.pptx` that PowerPoint itself can't open, because PowerPoint is the one writing it.

**🖼️ Export-to-verify** — the core differentiator. After any visual edit, export the slide (or the
whole deck) to an image with `export_slide_to_image` / `export_all_slides_to_images` and let a
vision-capable AI assistant *see* the result — catching overlapping shapes, text overflow, and
layout regressions that text-only automation simply cannot detect.

**Technical Requirements:**
- ⚠️ **Windows Only** — COM interop is Windows-specific
- ⚠️ **PowerPoint Required** — Microsoft PowerPoint 2016 or later must be installed
- ⚠️ **Desktop Environment** — controls a real PowerPoint process (not for server-side processing)

## 🎯 What You Can Do

**31 tools across 10 domains:**

- 🗂️ **Presentation** (5 tools) — create, open, save, close, list sessions
- 📑 **Slide** (3 tools) — add, count, delete
- ▭ **Shape** (6 tools) — rectangle, text box, count, delete, position, size
- ✏️ **TextFrame** (5 tools) — set/get text, font size, bold, font color
- 📊 **Table** (3 tools) — add, set/get cell text
- 🗣️ **Notes** (2 tools) — set/get speaker notes
- 🖼️ **Layout** (2 tools) — set/get slide layout
- 🖼️ **Image** (1 tool) — insert pictures
- 📈 **Chart** (2 tools) — add chart, get chart data
- 🖼️ **Export** (2 tools) — export a slide, or all slides, to images for visual verification

📚 **[Complete Feature Reference →](https://powerpointmcpserver.dev/features/)** — detailed
documentation of all 31 tools

## 💬 Example Prompts

**Build a deck from scratch:**
- *"Create a new presentation with a title slide and three content slides about our Q3 results,
  then export it as images so I can see it."*

**Tables &amp; charts:**
- *"Add a 4x3 table summarizing this data, then add a bar chart next to it."*

**Formatting &amp; shapes:**
- *"Make the title bold and blue, and move the logo to the top-right corner."*

**Speaker notes:**
- *"Write speaker notes for each slide summarizing the key talking point."*

**Visual verification:**
- *"Export slide 3 as an image and tell me if the chart overlaps the text box."*

## 👥 Who Should Use This?

**Perfect for:**
- ✅ AI assistants and coding agents that need to build or edit `.pptx` decks
- ✅ Anyone automating repetitive slide-deck workflows (reports, status decks, templates)
- ✅ Teams that want export-to-verify visual checks on every automated edit

**Not suitable for:**
- ❌ Server-side/headless processing (this drives a real desktop PowerPoint process)
- ❌ Linux/macOS users (Windows + PowerPoint installation required)

## 🚀 Quick Start

| Platform | Installation |
|----------|-------------|
| **VS Code** | [Install Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp) (one-click, recommended) |
| **Claude Desktop** | Download `.mcpb` from [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest) |
| **Any MCP Client** | Download `mcp-powerpoint.exe` from [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest) and add to PATH |
| **Details** | 📖 [Full Installation Guide →](https://powerpointmcpserver.dev/installation/) |

**⚠️ Important:** Close any open instances of the target file before automating it — the server
needs exclusive access to the presentation while it's driving PowerPoint.

## 🔧 CLI vs MCP Server

This project provides both a **CLI** and an **MCP Server** interface. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`pptcli` / `powerpointcli.exe`) | Coding agents (Copilot, Cursor, Windsurf) + scripting | Single tool, no large schemas — better for cost-sensitive, high-throughput automation. |
| **MCP Server** (`mcp-powerpoint`) | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent session. Better for interactive, exploratory workflows. |

**Manual installation:**
```powershell
# Primary: Download standalone executables from latest release (no .NET runtime required)
# https://github.com/sbroenne/mcp-server-powerpoint/releases/latest
# - PowerPointMcp-MCP-Server-{version}-windows.zip → extract mcp-powerpoint.exe
# - PowerPointMcp-CLI-{version}-windows.zip → extract powerpointcli.exe

# Secondary: Install via .NET tool (requires .NET 10 runtime)
dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
dotnet tool install --global Sbroenne.PowerPointMcp.CLI

# After installing either way, auto-configure your coding agents:
npx add-mcp "mcp-powerpoint" --name powerpoint-mcp
```

```powershell
# Optional: Install the agent skill for better AI guidance
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

> 💡 The VS Code extension installs this skill automatically. Manual `npx skills add` is for other
> MCP clients (Claude Code, Cursor, Windsurf, etc.).

## ⚙️ How It Works — Live COM Automation

**PowerPointMcp uses Windows COM automation to control the actual PowerPoint application (not
just `.pptx` files).**

```
┌──────────────────────┐        ┌──────────────────────┐
│  MCP Server          │        │  CLI (pptcli)        │
│  (AI assistants)     │        │  (coding agents)     │
└──────────┬───────────┘        └──────────┬───────────┘
           │ in-process                     │ named pipe →
           │ session registry               │ background daemon
           ▼                                ▼
┌──────────────────────┐        ┌──────────────────────┐
│  PowerPoint COM API  │        │  PowerPoint COM API  │
│  (PowerPoint.        │        │  (PowerPoint.        │
│   Application)       │        │   Application)       │
└──────────────────────┘        └──────────────────────┘
```

The MCP Server keeps sessions **in-process** in a `PresentationSessionRegistry` (no pipe, no
daemon) — sessions live as long as the MCP host process. The CLI instead keeps one PowerPoint
instance alive in a small **background daemon** (`pptcli service run`, auto-started on first use),
so a session persists across multiple separate `pptcli` invocations even though each command is
its own OS process.

**Key Benefits:**
- ✅ **True fidelity** — every render, export, and edit happens inside real PowerPoint, so what you
  get is exactly what PowerPoint would produce
- ✅ **Session-based workflow** — `open_presentation`/`create_presentation` start a session; every
  subsequent tool call operates on that session by `session_id`
- ✅ **Export-to-verify** — close the loop on every visual change with a real rendered image

## 📋 Additional Information

📚 **[MCP Server Guide →](src/PowerPointMcp.McpServer/README.md)** | **[CLI Guide →](src/PowerPointMcp.CLI/README.md)** | **[Agent Skills →](skills/README.md)**

📖 **[Complete Feature Reference](https://powerpointmcpserver.dev/features/)** •
**[Installation Guide](https://powerpointmcpserver.dev/installation/)** •
**[Changelog](https://powerpointmcpserver.dev/changelog/)** •
**[Contributing](https://powerpointmcpserver.dev/contributing/)** •
**[Security](https://powerpointmcpserver.dev/security/)** •
**[Privacy](https://powerpointmcpserver.dev/privacy/)**

**License:** MIT License — see [LICENSE](LICENSE)

**Built With:** This entire project was developed using GitHub Copilot AI assistance.

**Acknowledgments:**
- Microsoft PowerPoint Team — for comprehensive COM automation APIs
- Model Context Protocol community — for the AI integration standard
- [mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) — the sibling project this
  architecture and tooling is ported from

## Related Projects

Other projects by the author:

- [Excel MCP Server](https://excelmcpserver.dev/) — AI-powered Excel automation via Power Query,
  DAX, VBA, and PivotTables
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via MCP
- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered
  testing framework for AI agents
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio
  automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video
  generation

