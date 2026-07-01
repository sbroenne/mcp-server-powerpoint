# Ripley — History

## Seed (2026-07-01)

- **Project:** mcp-server-powerpoint — COM/PIA PowerPoint automation MCP server, C#/.NET 10, sibling to mcp-server-excel.
- **My domain:** `tests/PowerPointMcp.Core.Tests` — real-COM xUnit integration tests (no mocking), strict TDD (Rule 30).
- **Requested by:** sbroenne
- **Done so far:** Integration test files for all 9 Core domains (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart). The suite already caught two genuine COM bugs via TDD: `Presentations.Add()` creating zero slides, and `Presentations.Open(Untitled: msoTrue)` breaking `Save()`.
- **Test-fixture rule:** `xunit.runner.json` set to `parallelizeTestCollections: false`, `maxParallelThreads: 1` — without it, concurrent PowerPoint launches intermittently hit `COMException 0x800706BA "RPC server is unavailable"` (transient COM activation failure, not a product bug; same class mcp-server-excel documents).
- **Standing rules:** Integration tests only; real `.pptx` verified on disk; every test leaves no lingering `POWERPNT.exe`. Not yet covered: some error paths (locked file, bad path, timeout) and any Export/Open-Close domain once built.


📌 Team update (2026-07-01T11:07:03+02:00): Crash recovery completed; roadmap restored and Dallas test strategy assigns Ripley Phase 1 MCP validation: tools list/schema/error envelopes without COM where possible, session lifecycle with real PowerPoint only for wiring, serialized tests, and no duplicate Core behavior retests — decided by Squad-Coordinator/Dallas.


📌 Team update (2026-07-01T13:00:00+02:00): Phase 1 MVP landed green — In-memory MCP transport harness built (6/6 tests passed). Architecture: pipe pair → Program.ConfigureTestTransport → McpClient over StreamClientTransport, serialized via xunit.runner.json + [Collection("ProgramTransport")]. Two test classes: McpProtocolTests (5 tools validated, DI registry isolation confirmed, ServerInfo/Instructions green) + McpRoundTripTests (live-COM create_presentation disk write + full session lifecycle). Observation: POWERPNT.exe post-Quit takes ~90–100s to exit process list (documented Office COM cleanup, benign, self-resolves). Brett delivered McpServer MVP (5 tools, registry), Parker delivered Export (5/5 tests). Harness ready for reuse by future domain tool tests. — decided by Ripley
