# Dallas — History

## Seed (2026-07-01)

- **Project:** mcp-server-powerpoint — COM/PIA-based PowerPoint automation MCP server, C#/.NET 10, sibling to mcp-server-excel. Drives a live PowerPoint desktop instance via `Microsoft.Office.Interop.PowerPoint` (official PIA, embedded via `ForceEmbedPowerPointInteropTypes`). Windows + PowerPoint desktop required.
- **Architecture:** `ComInterop` (STA thread + OLE message filter + channel-based work queue, ported from ExcelBatch) → `Core` (domain commands) → `Service` (not built) → `CLI` / `MCP Server`.
- **Requested by:** sbroenne
- **State as of seeding:** 9 Core domains implemented with TDD real-COM integration tests (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart). `PowerPointMcp.McpServer` project folder exists but is EMPTY. No Service daemon, no Generators, no CI/hooks. Repo is local-only — no GitHub remote yet.
- **Key conventions:** Rule 1/1b (Success/ErrorMessage invariant, no try-catch suppression, validation vs exception-propagation distinction). Rule 30 (integration tests only, real COM, no mocking, strict TDD red→green). Tests serialized via `xunit.runner.json` `maxParallelThreads: 1` to avoid concurrent PowerPoint launches.
