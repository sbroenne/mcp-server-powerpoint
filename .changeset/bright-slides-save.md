---
"Sbroenne.PowerPointMcp.Core": minor
"Sbroenne.PowerPointMcp.ComInterop": minor
"Sbroenne.PowerPointMcp.Service": minor
"Sbroenne.PowerPointMcp.McpServer": minor
"Sbroenne.PowerPointMcp.CLI": minor
"Sbroenne.PowerPointMcp.Skill": minor
---

Add `save-as` and `save-copy-as` presentation operations to the MCP server and CLI. Save As moves
the active session to the new file only after PowerPoint succeeds, while Save Copy As preserves the
active file and session path. Both operations validate formats, paths, and explicit overwrite intent.
