---
name: powerpoint-cli
description: >
  Discover and use the PowerPoint CLI plugin to automate a live PowerPoint desktop session.
  Use for terminal, PowerShell, batch, and coding-agent tasks; not general presentation design.
  Triggers: pptcli, PowerPoint command line, presentation automation script.
compatibility: Windows with Microsoft PowerPoint desktop and Node.js 18 or later.
---

# PowerPoint CLI

Use the installed PowerPoint CLI plugin wrapper when available; otherwise use the installed
`pptcli` command. Start with `pptcli --help`, then `pptcli <group> --help` for the current
commands and options. The live help is the source of truth; this skill does not copy the command
catalog.

For editing, create or open one session and reuse the returned `sessionId` with later commands.
Slide, shape, row, and column indexes start at 1. Close with `session close --save` only when the
user wants to keep the changes. For visual edits, export images and inspect them before saving.

Use the optional `powerpoint-deck-design` skill only when creating or visually redesigning a
deck. See the [CLI guide](https://powerpointmcpserver.dev/cli/) and
[workflow guide](https://powerpointmcpserver.dev/reference/workflows/).
