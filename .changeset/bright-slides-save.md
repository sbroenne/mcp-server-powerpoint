---
"powerpointmcp": patch
---

Add `save-as` and `save-copy-as` presentation operations to the MCP server and CLI. Save As moves
the active session to the new file only after PowerPoint succeeds, while Save Copy As preserves the
active file and session path. Both operations validate formats, paths, and explicit overwrite intent.
