---
title: Claude Desktop Setup
description: Connect PowerPoint MCP Server to Claude Desktop on Windows — one-click MCPB bundle install or manual claude_desktop_config.json setup, with a copy-paste config block.
keywords: "Claude Desktop PowerPoint, Claude PowerPoint MCP, claude_desktop_config.json PowerPoint, MCPB bundle PowerPoint, Claude Desktop MCP Windows, automate PowerPoint with Claude"
---

# Using PowerPoint MCP Server with Claude Desktop

Connect Claude Desktop to a live Microsoft PowerPoint instance on Windows so
Claude can build and edit real `.pptx` decks — and *see* its own work by
exporting slides to images.

!!! info "Prerequisites"
    Windows, with Microsoft PowerPoint 2016 or later (desktop) installed and
    activated. See the [FAQ](faq.md#requirements-and-platform) for the full
    requirements.

## Option 1 — MCPB bundle (recommended)

The `.mcpb` bundle is a one-click install for Claude Desktop. It is
self-contained, so you do **not** need the .NET runtime.

1. Download the `.mcpb` file from the
   [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest).
2. Double-click it, or open it from Claude Desktop's extensions screen.
3. Confirm the install prompt.

Claude Desktop registers the server automatically — no config file editing.

## Option 2 — Manual configuration

If you prefer to manage the config yourself, first install the server:

=== "Standalone executable"

    Download `PowerPointMcp-MCP-Server-{version}-windows.zip` from the
    [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest),
    extract it, and note the full path to `mcp-powerpoint.exe`.

=== ".NET tool"

    Requires the .NET 10 runtime.

    ```powershell
    dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
    ```

    This puts `mcp-powerpoint` on your `PATH`.

Then open Claude Desktop → **Settings** → **Developer** → **Edit Config**,
which opens `claude_desktop_config.json`:

```text
%APPDATA%\Claude\claude_desktop_config.json
```

Add the `powerpoint` entry:

=== ".NET tool on PATH"

    ```json
    {
      "mcpServers": {
        "powerpoint": {
          "command": "mcp-powerpoint"
        }
      }
    }
    ```

=== "Standalone executable"

    ```json
    {
      "mcpServers": {
        "powerpoint": {
          "command": "C:\\Tools\\PowerPointMcp\\mcp-powerpoint.exe"
        }
      }
    }
    ```

!!! warning "Escape your backslashes"
    JSON requires `\\` for every backslash in a Windows path. A single `\` is
    the most common cause of a server that silently fails to start.

Restart Claude Desktop completely (quit from the system tray — closing the
window is not enough) for the change to take effect.

## Verify the connection

Claude Desktop shows connected MCP servers under the tools icon in the message
composer. You should see the 13 PowerPoint tools listed.

Then ask Claude:

> "Create a new PowerPoint presentation at `C:\Decks\demo.pptx` with a title
> slide, then export it as an image so I can see it."

Claude should call `presentation(action="create", ...)`, add content, and
finish with `export(action="export-slide-to-image", ...)`, returning a rendered
PNG of a real PowerPoint slide.

## Add the agent skill

The [agent skill](skills.md) gives Claude workflow guidance beyond the raw tool
schemas — session handling, export-to-verify loops and layout conventions:

```powershell
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| Server does not appear in Claude Desktop | Fully quit Claude Desktop from the system tray and reopen; check the JSON is valid and backslashes are escaped |
| "PowerPoint is not installed" | The full desktop app must be installed and activated — PowerPoint Online and Mobile cannot be automated |
| `mcp-powerpoint` not found | Restart your terminal *and* Claude Desktop after `dotnet tool install` so the updated `PATH` is picked up, or use the absolute `.exe` path |
| Edits fail on a file you have open | Close the target presentation first — the server needs exclusive access |

More detail in the [troubleshooting guide](troubleshooting.md).

## Next steps

- [Complete feature reference](features.md) — all 13 tools, 141 operations
- [MCP Server reference](mcp-server.md) — tool and session model
- [VS Code setup](vs-code.md) · [Cursor setup](cursor.md)
