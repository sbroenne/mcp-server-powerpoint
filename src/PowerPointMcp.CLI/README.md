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

pptcli <category> <action> --file <path-to.pptx> [options]
```

The first four commands are a minimal hand-written placeholder proving the
`ComInterop -> Core -> CLI` vertical slice end to end. Every other command
(`chart`, `image`, `layout`, `notes`, `presentation`, `shape`, `table`,
`textframe`) is source-generated from the same `[ServiceCategory]` Core
interfaces the MCP server uses — run `pptcli <category> --help` to see its
actions and options.

Each invocation is **fully self-contained**: it opens the target `.pptx` file,
runs exactly one action, saves, and closes — there is no persistent
session/daemon (PowerPointMcp deliberately has no out-of-process Service, so
`--file <path>` replaces what would otherwise be a `--session <id>`).

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
