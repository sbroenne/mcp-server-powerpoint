---
layout: default
title: "Automate PowerPoint with AI — Live COM Automation"
description: "Control Microsoft PowerPoint with natural language through AI assistants like GitHub Copilot and Claude. The only .NET-native MCP server that drives a live PowerPoint desktop instance, with export-to-verify multimodal checking."
keywords: "PowerPoint automation, MCP server, AI PowerPoint, slide automation, PowerPoint COM, GitHub Copilot PowerPoint, Claude PowerPoint, PowerPoint CLI, presentation automation"
canonical_url: "https://powerpointmcpserver.dev/"
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <img src="{{ '/assets/images/icon.png' | relative_url }}" alt="PowerPoint MCP Server Icon" class="hero-icon">
      <h1 class="hero-title">PowerPoint MCP Server</h1>
      <p class="hero-subtitle">Automate PowerPoint with AI via GitHub Copilot, Claude, and other MCP clients — driving a real, live PowerPoint desktop instance.</p>
    </div>
  </div>
</div>

<div class="badges-section">
  <div class="container">
    <div class="hero-badges">
      <a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp"><img src="https://img.shields.io/visual-studio-marketplace/i/sbroenne.powerpoint-mcp?label=VS%20Code%20Installs" alt="VS Code Marketplace Installs"></a>
      <a href="https://github.com/sbroenne/mcp-server-powerpoint"><img src="https://img.shields.io/github/stars/sbroenne/mcp-server-powerpoint?style=flat&label=GitHub%20Stars" alt="GitHub Stars"></a>
      <a href="https://github.com/sbroenne/mcp-server-powerpoint/releases"><img src="https://img.shields.io/github/downloads/sbroenne/mcp-server-powerpoint/total?label=GitHub%20Downloads" alt="GitHub Downloads"></a>
      <a href="https://www.nuget.org/packages/Sbroenne.PowerPointMcp.McpServer"><img src="https://img.shields.io/nuget/dt/Sbroenne.PowerPointMcp.McpServer.svg?label=NuGet%20MCP%20Downloads" alt="NuGet MCP Server Downloads"></a>
      <a href="https://www.nuget.org/packages/Sbroenne.PowerPointMcp.CLI"><img src="https://img.shields.io/nuget/dt/Sbroenne.PowerPointMcp.CLI.svg?label=NuGet%20CLI%20Downloads" alt="NuGet CLI Downloads"></a>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">
## 🤔 What is This?

**Automate PowerPoint with AI — a Model Context Protocol (MCP) server for live, real-time
PowerPoint automation through conversational AI.**

<div class="quick-install-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin: 24px 0;">
  <div style="text-align: center;">
    <p><strong>VS Code / GitHub Copilot</strong></p>
    <a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp" class="button-link">Install Extension</a>
  </div>
  <div style="text-align: center;">
    <p><strong>Claude Desktop</strong></p>
    <a href="https://github.com/sbroenne/mcp-server-powerpoint/releases/latest" class="button-link">One-Click Install (MCPB)</a>
  </div>
  <div style="text-align: center;">
    <p><strong>NuGet .NET Tool</strong></p>
    <a href="/installation/" class="button-link">dotnet tool install</a>
  </div>
  <div style="text-align: center;">
    <p><strong>Cursor, Windsurf, etc.</strong></p>
    <a href="/installation/" class="button-link">Installation Guide</a>
  </div>
</div>

**MCP Server for PowerPoint** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to build
and edit real `.pptx` presentations through natural language — slides, shapes, text, tables,
charts, speaker notes, and layouts — no VBA or PowerPoint object-model knowledge required.

It works with any MCP-compatible AI assistant like GitHub Copilot, Claude Desktop, Cursor,
Windsurf, etc.

**🛡️ Why This Is Different — Live COM Automation, Not File Parsing** — Most PowerPoint MCP
servers manipulate `.pptx` files offline with libraries like `python-pptx`, or use agent-run
scripts with LibreOffice-rendered thumbnails. This project instead drives a **live, real
PowerPoint desktop instance** via `Microsoft.Office.Interop.PowerPoint` — the official Primary
Interop Assembly. That means true-fidelity rendering, compatibility with an already-open deck,
and zero risk of producing a `.pptx` that PowerPoint itself can't open, because PowerPoint is the
one writing it.

**🖼️ Export-to-Verify** — the core differentiator. After any visual edit, export the slide (or
the whole deck) to an image with `export_slide_to_image` / `export_all_slides_to_images` and let
a vision-capable AI assistant *see* the result — catching overlapping shapes, text overflow, and
layout regressions that text-only automation simply cannot detect.

## Key Features

<div class="features-grid">
<div class="feature-card">
<h3>Slides &amp; Layouts</h3>
<p>Add and delete slides, apply and inspect layouts, and query slide count for state discovery.</p>
</div>

<div class="feature-card">
<h3>Shapes &amp; Text</h3>
<p>Add rectangles and text boxes, position and resize shapes, set and read rich text with font size, bold, and color.</p>
</div>

<div class="feature-card">
<h3>Tables &amp; Charts</h3>
<p>Build tables cell-by-cell and add charts with real data — then read the data back to verify.</p>
</div>

<div class="feature-card">
<h3>Speaker Notes</h3>
<p>Set and read presenter notes per slide for talk-track generation and review.</p>
</div>

<div class="feature-card">
<h3>Images</h3>
<p>Insert pictures from local files directly onto any slide.</p>
</div>

<div class="feature-card">
<h3>🖼️ Export-to-Verify</h3>
<p>Export any slide — or the whole deck — to images for multimodal visual verification. The project's core differentiator over text-only PowerPoint tooling.</p>
</div>
</div>

<p><a href="/features/">See all 31 tools across 10 domains →</a></p>

## What Can You Do With It?

Ask your AI assistant to build and edit presentations using natural language:

<div class="example-section">
<h4>📝 Build a Deck From Scratch</h4>
<p><strong>You:</strong> "Create a new presentation with a title slide and three content slides about our Q3 results, then export it as images so I can see it."</p>
<p>AI creates the presentation, adds slides with headings and body text, and exports PNGs of every slide to verify the result.</p>
</div>

<div class="example-section">
<h4>📊 Tables &amp; Charts</h4>
<p><strong>You:</strong> "Add a 4x3 table summarizing this data, then add a bar chart next to it."</p>
<p>AI builds the table cell-by-cell and adds a chart shape with the given data, then exports the slide to confirm the layout looks right.</p>
</div>

<div class="example-section">
<h4>🎨 Formatting &amp; Shapes</h4>
<p><strong>You:</strong> "Make the title bold and blue, and move the logo to the top-right corner."</p>
<p>AI applies text formatting through the TextFrame tools and repositions the shape, then exports an image to verify nothing overlaps.</p>
</div>

<div class="example-section">
<h4>🗣️ Speaker Notes</h4>
<p><strong>You:</strong> "Write speaker notes for each slide summarizing the key talking point."</p>
<p>AI reads each slide's content and writes tailored notes via <code>set_notes_text</code>.</p>
</div>

<div class="example-section">
<h4>🖼️ Visual Verification</h4>
<p><strong>You:</strong> "Export slide 3 as an image and tell me if the chart overlaps the text box."</p>
<p>AI exports the slide with <code>export_slide_to_image</code> and inspects the rendered PNG directly — catching issues no text-only tool could see.</p>
</div>

## CLI vs MCP Server

This project provides both a **CLI** and an **MCP Server** interface. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`pptcli`) | Coding agents (Copilot, Cursor, Windsurf) | Single tool, no large schemas — better for cost-sensitive, high-throughput automation. |
| **MCP Server** (`mcp-powerpoint`) | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent session. Better for interactive, exploratory workflows. |

## ⚙️ How It Works — Live COM Automation

**PowerPoint MCP Server uses Windows COM automation to control the actual PowerPoint application
(not just `.pptx` files).**

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server         │     │   CLI (pptcli)       │
│  (AI assistants)     │     │  (coding agents)     │
└─────────┬────────────┘     └─────────┬────────────┘
          │                            │
          └──────────┬─────────────────┘
                     ▼
          ┌─────────────────────────┐
          │  Session Registry       │
          │ (in-process, per-server)│
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   PowerPoint COM API    │
          │ (PowerPoint.Application)│
          └─────────────────────────┘
```

**Key Benefits:**
- ✅ **True Fidelity** — every render, export, and edit happens inside real PowerPoint, so what
  you get is exactly what PowerPoint would produce
- ✅ **Session-Based Workflow** — `open_presentation`/`create_presentation` start a session;
  every subsequent tool call operates on that session by `session_id`
- ✅ **Export-to-Verify** — close the loop on every visual change with a real rendered image

## Documentation

📖 **[Complete Feature Reference](/features/)** — All 31 tools across 10 domains

📥 **[Installation Guide](/installation/)** — Setup for VS Code, Claude Desktop, other MCP clients, and CLI

📖 **[MCP Server Documentation](/mcp-server/)** — Complete MCP tool reference and examples

📖 **[CLI Documentation](/cli/)** — Full CLI command reference and examples

🤖 **[Agent Skills](/skills/)** — Cross-platform AI guidance (auto-installed by the VS Code extension)

📋 **[Changelog](/changelog/)** — Release notes and version history

## More Information

- [GitHub Repository](https://github.com/sbroenne/mcp-server-powerpoint) — Source code, issues, and contributions
- [Contributing Guide](/contributing/) — How to contribute

## Related Projects

Other projects by the author:

- [Excel MCP Server](https://excelmcpserver.dev/) — the sibling project this port is based on: AI-powered Excel automation via Power Query, DAX, VBA, and PivotTables
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via GitHub Copilot, Claude, and other MCP clients
- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered testing framework for AI agents
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation

<footer>
<div class="container">
<p><strong>PowerPoint MCP Server</strong> — MIT License — © 2025-2026</p>
</div>
</footer>
</div>
