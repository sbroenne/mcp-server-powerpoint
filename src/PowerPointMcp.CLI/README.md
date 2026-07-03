# PowerPoint MCP — CLI

A command-line companion to the [PowerPoint MCP Server](https://www.nuget.org/packages/Sbroenne.PowerPointMcp.McpServer)
for scripting PowerPoint automation directly from a terminal, built on the same
`ComInterop → Core` engine.

> **Windows only.** Requires Microsoft PowerPoint (desktop) to be installed.

## Install

```pwsh
dotnet tool install --global Sbroenne.PowerPointMcp.CLI
```

This installs the `pptcli` command.

## Usage

```pwsh
pptcli create <path-to-new.pptx>
pptcli slide add-blank <path-to.pptx>
pptcli slide count <path-to.pptx>
pptcli slide delete <path-to.pptx> <slideIndex>
```

> The CLI currently exposes a focused command set proving the end-to-end
> automation pipeline; the full surface (shapes, text, tables, charts, images,
> notes, layouts, export) is available today via the MCP server. See the
> [documentation](https://powerpointmcpserver.dev) for the roadmap.

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
