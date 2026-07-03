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

📌 Brett (2026-07-01T13:10:00+02:00): Hand-wrote the 9 remaining domain tool classes — SlideTools, ShapeTools, TextFrameTools, TableTools, NotesTools, LayoutTools, ImageTools, ChartTools, ExportTools (26 new tools). MCP surface now 31 tools total (up from 5). Followed PresentationTools.cs pattern exactly: static `[McpServerToolType]` classes, static readonly `{Domain}Commands` instance, `registry.TryGet(sessionId, out batch)` guard, SerializeResult mapping `{Domain}OperationResult` → camelCase JSON. Chart's `IReadOnlyList<string>`/`<double>` params exposed as `string[]`/`double[]` — MCP SDK 1.3.0 serializes to JSON array schema with no extra mapping. Export's optional params (format/width/height with C# defaults) forced `registry` to `PresentationSessionRegistry? registry = null` (C# forbids required-after-optional in a signature) — DI resolution still works regardless of nullability, verified via tests/list. Build green (0 warn/0 err); ran subset of Ripley's McpProtocolTests (no-registry-leak, all-have-name-and-description, server-info) — all pass; the hardcoded 5-tool assertion now fails as expected (flagged for Ripley to update). Decision recorded to .squad/decisions/inbox/brett-remaining-domain-tools.md. — decided by Brett

📌 Team update (2026-07-02T14-15-00+02:00): Phase 1 complete — Brett's McpServer MVP (31 tools, async-close pattern, in-process registry) now part of consolidated decision. Parker's resilient shutdown hardening merged; create_presentation blocking finding flagged for future UX fix. Full session summary in decisions.md + orchestration-log/2026-07-02T14-15-00-brett.md — decided by Scribe

