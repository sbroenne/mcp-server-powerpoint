---
title: Cursor Setup
description: Connect PowerPoint MCP Server to Cursor on Windows — .cursor/mcp.json configuration, the pptcli CLI path for agent workflows, and a copy-paste config block.
keywords: "Cursor PowerPoint MCP, .cursor/mcp.json PowerPoint, Cursor MCP server Windows, automate PowerPoint with Cursor, Cursor agent PowerPoint, pptcli Cursor"
---

# Using PowerPoint MCP Server with Cursor

Connect Cursor to a live Microsoft PowerPoint instance on Windows so its agent
can build and edit real `.pptx` decks — and export slides to images to verify
the result visually.

!!! info "Prerequisites"
    Windows, with Microsoft PowerPoint 2016 or later (desktop) installed and
    activated. See the [FAQ](faq.md#requirements-and-platform) for the full
    requirements.

## Install the server

=== ".NET tool"

    Requires the .NET 10 runtime.

    ```powershell
    dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
    ```

    This puts `mcp-powerpoint` on your `PATH`.

=== "Standalone executable"

    Download `PowerPointMcp-MCP-Server-{version}-windows.zip` from the
    [latest release](https://github.com/sbroenne/mcp-server-powerpoint/releases/latest),
    extract it, and note the path to `mcp-powerpoint.exe`. No .NET runtime
    needed.

## Configure Cursor

The fastest route:

```powershell
npx add-mcp "mcp-powerpoint" --name powerpoint-mcp
```

Or configure it by hand. Cursor reads MCP servers from either location:

| Scope | Path |
|---|---|
| Project | `.cursor\mcp.json` in the repo root |
| Global | `%USERPROFILE%\.cursor\mcp.json` |

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

Open **Settings → MCP** in Cursor to confirm the `powerpoint` server shows as
connected with its tools listed.

## Prefer the CLI for agent workflows

Cursor's agent is a coding agent, and coding agents generally do better with the
CLI than with a large MCP tool surface — one command instead of 13 tool schemas
in every request, which is meaningfully cheaper at high throughput:

```powershell
dotnet tool install --global Sbroenne.PowerPointMcp.CLI
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

The [agent skill](skills.md) teaches Cursor the session model and the
export-to-verify loop. The `pptcli` daemon keeps presentations open across
invocations, so a multi-step script does not pay PowerPoint's start-up cost
every command. See the [CLI reference](cli.md).

## Verify the connection

Ask Cursor's agent:

> "Create a new PowerPoint presentation at `C:\Decks\demo.pptx` with a title
> slide, then export it as an image so I can see it."

It should call `presentation(action="create", ...)`, add content, and finish
with `export(action="export-slide-to-image", ...)`, returning a rendered PNG of
a real PowerPoint slide.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Server shows as disconnected in Settings → MCP | Check the JSON is valid and backslashes are escaped; restart Cursor |
| `mcp-powerpoint` not found | Restart Cursor after `dotnet tool install` so the updated `PATH` is inherited, or use the absolute `.exe` path |
| "PowerPoint is not installed" | The full desktop app must be installed and activated — PowerPoint Online and Mobile cannot be automated |
| Edits fail on a file you have open | Close the target presentation first — the server needs exclusive access |

More detail in the [troubleshooting guide](troubleshooting.md).

## Next steps

- [Complete feature reference](features.md) — all 13 tools, 141 operations
- [CLI reference](cli.md) — every `pptcli` command
- [Claude Desktop setup](claude-desktop.md) · [VS Code setup](vs-code.md)
