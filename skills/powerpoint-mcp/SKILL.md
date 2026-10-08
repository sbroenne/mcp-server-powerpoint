---
name: powerpoint-mcp
description: >
  Use the configured PowerPoint MCP server to inspect or edit presentations on Windows.
  Covers session safety, live tool discovery, saving, and visual verification; not general
  presentation design. Triggers: PowerPoint MCP, presentation tools, edit a PowerPoint deck.
compatibility: Windows with Microsoft PowerPoint desktop installed and the PowerPoint MCP server configured.
---

# PowerPoint MCP

Use this skill only for tasks performed through the configured PowerPoint MCP server. For
general presentation design, use the optional `powerpoint-deck-design` skill. Tool schemas
advertised by the live server define the available actions and arguments; do not rely on a
copied command catalog.

## Safe editing loop

1. Check the live tool descriptions. Use a `{domain}_read` tool when it provides the inspection
   action you need.
2. Create or open a presentation once. Pass the returned `presentation_session_id` to later
   calls; do not open a file a second time after `create`.
3. Indices for slides, shapes, rows, and columns start at 1.
4. Export and inspect images after visual changes. Before delivery, run the accessibility audit
   and close with `save: true` when the user asked to keep the changes.
5. Close returns before PowerPoint finishes its background cleanup; do not wait for its process
   to exit.

The server requires Windows and desktop PowerPoint; it is not a headless or cross-platform
converter. See the [workflow guide](https://powerpointmcpserver.dev/reference/workflows/) and
[behavioral rules](https://powerpointmcpserver.dev/reference/behavioral-rules/) for edge cases.
