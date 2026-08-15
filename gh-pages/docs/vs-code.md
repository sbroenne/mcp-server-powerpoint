---
title: VS Code & GitHub Copilot Setup
description: Connect PowerPoint MCP Server to GitHub Copilot Chat in VS Code — one-click extension install or manual mcp.json configuration, plus the pptcli CLI for coding agents.
keywords: "VS Code PowerPoint MCP, GitHub Copilot PowerPoint, mcp.json PowerPoint, VS Code MCP server Windows, Copilot Chat PowerPoint automation, pptcli VS Code"
---

# Using PowerPoint MCP Server with VS Code and GitHub Copilot

Give GitHub Copilot Chat in VS Code control of a live Microsoft PowerPoint
instance — building decks, editing slides, and exporting slides to images so
Copilot can visually verify its own work.

!!! info "Prerequisites"
    Windows, with Microsoft PowerPoint 2016 or later (desktop) installed and
    activated. See the [FAQ](faq.md#requirements-and-platform) for the full
    requirements.

## Option 1 — VS Code extension (recommended)

The extension registers the MCP server *and* installs the
[agent skill](skills.md) automatically to `~/.copilot/skills/powerpoint-mcp/`.

[Install from the Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.powerpoint-mcp){ .md-button .md-button--primary }

After installing, reload VS Code and open Copilot Chat in **Agent** mode. The
PowerPoint tools appear in the tools picker.

## Option 2 — Manual `mcp.json` configuration

First install the server:

=== ".NET tool"

    Requires the .NET 10 runtime.

    ```powershell
    dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
    ```

=== "Standalone executable"

    Download `PowerPointMcp-MCP-Server-{version}-windows.zip` from the
    [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest),
    extract it, and note the path to `mcp-powerpoint.exe`.

Then register it. The quickest way:

```powershell
npx add-mcp "mcp-powerpoint" --name powerpoint-mcp
```

Or edit the config by hand. Run **MCP: Open User Configuration** from the
Command Palette (`Ctrl+Shift+P`) to open your user `mcp.json`, or create
`.vscode/mcp.json` in a workspace to scope the server to one project:

=== ".NET tool on PATH"

    ```json
    {
      "servers": {
        "powerpoint": {
          "type": "stdio",
          "command": "mcp-powerpoint"
        }
      }
    }
    ```

=== "Standalone executable"

    ```json
    {
      "servers": {
        "powerpoint": {
          "type": "stdio",
          "command": "C:\\Tools\\PowerPointMcp\\mcp-powerpoint.exe"
        }
      }
    }
    ```

!!! warning "Escape your backslashes"
    JSON requires `\\` for every backslash in a Windows path. A single `\` is
    the most common cause of a server that silently fails to start.

## Option 3 — The CLI, for coding agents

If you mostly use Copilot in agent mode to run commands, the CLI is usually the
better fit: one tool instead of 13 schemas, which is meaningfully cheaper at
high throughput.

```powershell
dotnet tool install --global Sbroenne.PowerPointMcp.CLI
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

The `pptcli` daemon keeps presentations open across invocations, so a script
does not pay PowerPoint's start-up cost on every command. See the
[CLI reference](cli.md).

## Verify the connection

In Copilot Chat (Agent mode), ask:

> "Create a new PowerPoint presentation at `C:\Decks\demo.pptx` with a title
> slide, then export it as an image so I can see it."

Copilot should call `presentation(action="create", ...)`, add content, and
finish with `export(action="export-slide-to-image", ...)`, returning a rendered
PNG of a real PowerPoint slide.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Tools do not appear in Copilot Chat | Make sure you are in **Agent** mode, then reload the window; check the MCP server's output channel for start-up errors |
| `mcp-powerpoint` not found | Restart VS Code after `dotnet tool install` so the updated `PATH` is inherited, or use the absolute `.exe` path |
| "PowerPoint is not installed" | The full desktop app must be installed and activated — PowerPoint Online and Mobile cannot be automated |
| Edits fail on a file you have open | Close the target presentation first — the server needs exclusive access |

More detail in the [troubleshooting guide](troubleshooting.md).

## Next steps

- [Complete feature reference](features.md) — all 13 tools, 141 operations
- [CLI reference](cli.md) — every `pptcli` command
- [Claude Desktop setup](claude-desktop.md) · [Cursor setup](cursor.md)
