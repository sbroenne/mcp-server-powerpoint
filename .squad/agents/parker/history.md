# Parker — History

## Seed (2026-07-01)

- **Project:** mcp-server-powerpoint — COM/PIA PowerPoint automation MCP server, C#/.NET 10, sibling to mcp-server-excel. Drives a live PowerPoint desktop instance via the official PIA (embedded, `ForceEmbedPowerPointInteropTypes`).
- **My domain:** `ComInterop` (STA thread + OLE message filter + channel work queue, ported from ExcelBatch; simplified single-presentation) and `Core` domain commands.
- **Requested by:** sbroenne
- **Done so far:** 9 Core domains with real-COM TDD tests (Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart). Two real COM bugs already found+fixed: (1) `Presentations.Add()` creates ZERO slides — fixed by adding one blank slide before `SaveAs`; (2) `Presentations.Open(Untitled: msoTrue)` breaks `Save()` — fixed by passing `Untitled: msoFalse`.
- **Not done:** `Open`/`Close`/session-list on Presentation, `Export` domain (folder exists, empty), resilient close/quit retry (current shutdown is bare-bones, not the exponential-backoff ExcelShutdownService equivalent), multi-presentation batches, IRM/AIP, `.pptm` macro-security.
- **Conventions:** Rule 1/1b invariants, Rule 30 (real-COM integration tests only, strict TDD). Watch untyped `MsoTriState` int constants at every COM call site.


📌 Team update (2026-07-01T11:07:03+02:00): Crash recovery completed; roadmap restored and Dallas recommended MCP MVP starts with Excel-style stdio host, in-process session registry, and hand-written tools before generators. Parker Phase 1 owns Core/ComInterop follow-up: add Presentation Open/Close/session-list and Export domain after Brett's server vertical slice needs are clear — decided by Squad-Coordinator/Dallas.


📌 Team update (2026-07-01T13:00:00+02:00): Phase 1 MVP landed green — Export domain implemented with ExportSlideToImage + ExportAllSlidesToImages (5/5 real-COM integration tests passed). Key decisions: Presentation.Export for all-slides (single COM call, efficient), 1-based slide indices (matches PowerPoint native + other Core domains), format string passed directly to FilterName (PNG/JPG/GIF/BMP/TIF/WMF/EMF, no pre-validation per Rule 1b), output directory pre-created gracefully. Brett delivered McpServer (5 tools, registry, shutdown), Ripley built transport harness (6/6 tests). Export ready for MCP tool integration in Phase 2. — decided by Parker
