# Brett — History

## Seed (2026-07-01)

- **Project:** mcp-server-powerpoint — COM/PIA PowerPoint automation MCP server, C#/.NET 10, sibling to mcp-server-excel.
- **My domain:** MCP server, Service daemon, CLI, and source generators — the layers above Core.
- **Requested by:** sbroenne
- **State as of seeding:** `src/PowerPointMcp.McpServer` project folder exists but has ZERO `.cs` files — no MCP server yet. This is the biggest current gap now that 9 Core domains are solid. CLI is a minimal hand-written `Program.cs` placeholder (not the target Generators-based CLI). No `PowerPointMcp.Service`, no `Generators`/`.Cli`/`.Mcp`.
- **Guidance:** Copy+adapt `ExcelMcp.Generators*` and the Service/hosting pattern from mcp-server-excel. Do NOT invest in generators until the Core interface shape is confirmed stable (it now is, across 9 domains). Keep MCP/CLI layers thin — marshal to Core, no domain logic.


📌 Team update (2026-07-01T11:07:03+02:00): Crash recovery completed; roadmap restored and Dallas architecture plan makes Brett Phase 1 owner of the empty McpServer: csproj/Program stdio host, in-process PresentationSessionRegistry, PowerPointToolsBase, PresentationTool vertical slice, Slide/Shape tools, remaining domain tools, manifest/README — decided by Squad-Coordinator/Dallas.


📌 Team update (2026-07-01T11:40:00+02:00): Scaffolding + decisions now on origin/main. Coordinator merged 2 decisions (release-deliverables, GitHub Pages domain), staged 204 framework files, committed 655e867, pushed main → origin/main. Framework ready for team work — decided by Coordinator


📌 Team update (2026-07-01T13:00:00+02:00): Phase 1 MVP landed green — McpServer stdio host built with 5 Presentation tools (create/open/save/close/list_sessions), in-process PresentationSessionRegistry with two-layer shutdown disposal, PowerPointToolsBase (JSON/error serialization), hand-written tools pattern (DI injection into static methods, registry correctly isolated from MCP schema). ModelContextProtocol bumped 0.5.0-preview.2 → 1.3.0. Full solution builds green (0 warn/0 err). Parker delivered Export domain (5/5 real-COM tests passed). Ripley built transport harness (6/6 protocol + live-COM tests). Remaining 8 domain tools (Slide/Shape/TextFrame/Table/Notes/Layout/Image/Chart) ready for same hand-written pattern; generators deferred to Phase 2. — decided by Brett
