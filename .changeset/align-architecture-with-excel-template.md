---
"powerpointmcp": minor
---

Aligned mcp-server-powerpoint's architecture and tooling with mcp-server-excel, the
authoritative template for this family of projects:

- **Single MCP dispatch tool**: replaced the 12 separate presentation-related MCP tools
  with one `presentation` tool taking an `action` parameter, matching Excel's dispatch
  pattern. Added matching CLI session-management commands so the CLI and MCP Server
  remain equal, parity-checked entry points.
- **COM object cleanup**: every `dynamic` COM object obtained inside a Core Commands
  method is now released in a `finally` block once no longer needed, instead of relying
  solely on the top-level session Application/Presentation cleanup. This closes a class
  of potential COM object leaks under heavy or long-running use.
- **Crash-safety process tracking**: PowerPoint process IDs are now tracked in a
  process-wide registry with an `AppDomain.ProcessExit` handler that force-kills any
  still-running `POWERPNT.exe` if the host process itself terminates uncleanly (crash,
  forced kill, etc.), closing a gap versus Excel's existing crash-safety net.
- **Ported audit scripts** from mcp-server-excel: `check-com-leaks.ps1`,
  `check-success-flag.ps1`, `check-dynamic-casts.ps1`, and a new
  `check-core-interface-completeness.ps1` tailored to this project's
  generate-enums-from-interface architecture. All are wired into the pre-commit hook.
- Corrected stale documentation (tool/operation/domain counts) across the README,
  docs site, and skills files.

No behavior change for existing MCP tool calls beyond the dispatch-tool consolidation;
CLI commands are unaffected.
