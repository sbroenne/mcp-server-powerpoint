# MCP Server for PowerPoint

<!-- mcp-name: io.github.sbroenne/mcp-server-powerpoint -->
mcp-name: io.github.sbroenne/mcp-server-powerpoint

Let your AI assistant create, edit, and export Microsoft PowerPoint presentations.
Build slides, format text and shapes, add charts and tables, insert images, and
write speaker notes. Export PDFs for sharing or slide images so your assistant
can visually review the result.

This [Model Context Protocol](https://modelcontextprotocol.io) server controls
the installed PowerPoint desktop application on Windows.

> **Windows only.** Requires Microsoft PowerPoint (desktop) to be installed.

## Install

```pwsh
dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
```

This installs the `mcp-powerpoint` command. Point your MCP client (Claude
Desktop, VS Code, GitHub Copilot, etc.) at it over stdio:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "mcp-powerpoint"
    }
  }
}
```

## Capabilities

**32 tools with 215 operations across 17 domains:**

| Tool | Ops | Coverage |
| --- | --- | --- |
| `presentation` | 20 | create, open, Save As, Save Copy As, close, list, templates, get/set final, document properties, string tags |
| `slide` | 26 | slide lifecycle, visibility, backgrounds, sections, comments, import, string tags |
| `shape` | 57 | shape creation, free-floating and attached connectors, editable WordArt, per-axis 3D rotation, alignment/distribution, styling, grouping, boolean merging, hyperlinks, placeholders, string tags, `copy-formatting`, `duplicate`, `copy-to-slide`, `get-link-info`, `update-link`, `break-link`, `set-link-auto-update` |
| `textframe` | 22 | text content, find/replace, font formatting, alignment, bullets, auto-size |
| `table` | 12 | tables, cell text, row/column edits, cell fill/border, merge |
| `notes` | 2 | set/get speaker notes |
| `layout` | 4 | set/get slide layout |
| `pagesetup` | 5 | slide size, numbering, footer, date/time |
| `accessibility` | 3 | deterministic audit and reading order |
| `master` | 12 | theme color/font inspection, title/body placeholder fonts, solid + gradient master backgrounds |
| `animation` | 5 | shape effects, transition read/write |
| `image` | 16 | add picture, brightness/contrast, recolor, crop |
| `media` | 2 | add embedded/linked audio or video, get native media info |
| `chart` | 16 | charts, series, titles, legend, data replacement, styles, colors, data tables |
| `smartart` | 7 | SmartArt insertion and node editing |
| `export` | 3 | export to PDF or images |
| `customshow` | 3 | named custom shows: curated, ordered subsets of a presentation's slides |

Every domain is exposed as a single **action-dispatch tool** taking an `action`
parameter — including `presentation`. Example MCP calls:

- `presentation(action="create", filePath="C:\\Decks\\demo.pptx")`
- `presentation(action="apply-template", presentation_session_id="...", templatePath="C:\\Templates\\brand.potx")`
- `chart(action="add-chart", presentation_session_id="...", slide_index=2, ...)`

## How it works

`ComInterop` (STA thread + OLE message filter) → `Core` (domain command classes) →
`PowerPointMcp.Service` (shared service layer) → `McpServer` (stdio host, calling the service
in-process) and `CLI` (talks to the same service via a background named-pipe daemon). Each open
presentation is a session identified by a `presentation_session_id`; tools operate on a session until it is
closed.

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
