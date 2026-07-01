# Dallas — History

## Seed (2026-07-01)

- **Project:** mcp-server-powerpoint — COM/PIA-based PowerPoint automation MCP server, C#/.NET 10, sibling to mcp-server-excel. Drives a live PowerPoint desktop instance via `Microsoft.Office.Interop.PowerPoint` (official PIA, embedded via `ForceEmbedPowerPointInteropTypes`). Windows + PowerPoint desktop required.
- **Architecture:** `ComInterop` (STA thread + OLE message filter + channel-based work queue, ported from ExcelBatch) → `Core` (domain commands) → `Service` (not built) → `CLI` / `MCP Server`.
- **Requested by:** sbroenne
- **State as of seeding:** 9 Core domains implemented with TDD real-COM integration tests (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart). `PowerPointMcp.McpServer` project folder exists but is EMPTY. No Service daemon, no Generators, no CI/hooks. Repo is local-only — no GitHub remote yet.
- **Key conventions:** Rule 1/1b (Success/ErrorMessage invariant, no try-catch suppression, validation vs exception-propagation distinction). Rule 30 (integration tests only, real COM, no mocking, strict TDD red→green). Tests serialized via `xunit.runner.json` `maxParallelThreads: 1` to avoid concurrent PowerPoint launches.


📌 Team update (2026-07-01T11:40:00+02:00): Scaffolding + decisions now on origin/main. Coordinator merged 2 decisions (release-deliverables, GitHub Pages domain), staged 204 framework files, committed 655e867, pushed main → origin/main. Framework ready for team work — decided by Coordinator


📌 Team update (2026-07-01T13:00:00+02:00): Phase 1 MVP landed green across all three concurrent work streams. Brett: McpServer stdio host (5 tools, in-process registry, DI injection pattern, two-layer shutdown, full solution green). Parker: Export domain (ExportSlideToImage + ExportAllSlidesToImages, 5/5 real-COM tests). Ripley: MCP transport harness (6/6 protocol + live-COM tests, reusable for future domains). Vertical slice proven (session → registry → Core → MCP serialization). All three inbox decisions consolidated into decisions.md. Architecture pass validated against real repo state; 1.3.0 SDK fluent API confirmed; generators deferred to Phase 2 (Core shape now stable). Office post-Quit ~90s latency documented (benign). Ready for Phase 2: remaining 8 domain tools, generators port, out-of-process Service. — decided by Scribe
