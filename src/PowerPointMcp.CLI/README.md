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

## How it works — a background daemon, not a cold start per command

Launching PowerPoint and tearing it down again can take 90-150 seconds. To avoid paying that
cost on every single CLI invocation, `pptcli` keeps one PowerPoint instance alive in a small
background daemon (`pptcli service run`) that is auto-started the first time you open or create
a session, and reused by every subsequent command that references the same session id — even
though each command is a separate OS process. This mirrors the architecture used by
`mcp-server-excel`'s CLI.

```pwsh
# Create a new presentation — this may auto-start the daemon (first PowerPoint launch cost).
pptcli session create C:\decks\demo.pptx
# {"success":true,"sessionId":"…","filePath":"C:\\decks\\demo.pptx"}

# Every subsequent command reuses the SAME PowerPoint process via --session/-s.
pptcli slide add-blank -s <SESSION_ID>
pptcli slide get-count -s <SESSION_ID>
pptcli notes set-notes-text -s <SESSION_ID> --slide-index 1 --text "Speaker notes"

# Save and close when you're done with the file (closing does not stop the daemon).
pptcli session close <SESSION_ID> --save

# Check on / manage the daemon directly.
pptcli service status
pptcli service stop
pptcli service stop --force   # stop only PID + start-time identities owned by this daemon
```

Every `[ServiceCategory]`-annotated Core domain (Presentation, Slide, Shape, TextFrame, Table,
Chart, Image, Media, Notes, Layout) gets a generated `pptcli <category> <action> -s <SESSION_ID> [options]`
command automatically — run `pptcli <category> --help` to see the actions and options available
for that domain. (Export is not yet exposed via the CLI — a known, pre-existing generator
conflict, unrelated to session handling.)

The daemon shuts down automatically after an idle period with no open sessions, or immediately
via `pptcli service stop`.

Each pipe has an atomic ownership record containing the daemon and PowerPoint process PID plus
start time. Forced cleanup validates both values and refuses to terminate anything it cannot prove
belongs to that daemon. `service status` reports `starting`, `running`, `unresponsive`, or
`stopped` rather than treating every failed ping as stopped.

## Session commands

| Command | Description |
|---|---|
| `pptcli session create <path>` | Create a new presentation, return a session id. |
| `pptcli session open <path>` | Open an existing presentation, return a session id. |
| `pptcli session test <path>` | Validate that PowerPoint can open a presentation without retaining a session. |
| `pptcli session close <id> [--save]` | Close a session, optionally saving first. |
| `pptcli session list` | List every session currently open in the daemon. |

## Service (daemon) commands

| Command | Description |
|---|---|
| `pptcli service start` | Auto-start the daemon if it isn't already running. |
| `pptcli service status` | Report the daemon state, responsiveness, session count, and uptime. |
| `pptcli service stop [--force]` | Gracefully shut down the daemon, or stop only validated owned processes. |

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
