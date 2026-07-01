# Continuation Notes (Session 1)

This file tracks exactly what was scaffolded in the first working session vs. what remains,
so a follow-up session (or a human) can pick this up without re-deriving context.

## Update: real PowerPoint COM validation (Session 1, same session)

This machine actually has PowerPoint installed, so the vertical slice was validated for
real, not just build-checked:

```powershell
dotnet run --project src\PowerPointMcp.CLI -- create C:\temp\demo.pptx
# => Created: C:\temp\demo.pptx (34,237 bytes, verified on disk)
```

Confirmed: no lingering `POWERPNT.exe` process after the batch disposed — the STA
thread + `PresentationShutdownService.CloseAndQuit` cleanup path works correctly on a
real PowerPoint installation. This meaningfully de-risks the port: the
`Presentations.Add`/`SaveAs`/`Close`/`Quit` COM calls and the late-bound `MsoTriState`
workarounds are confirmed correct against real PowerPoint, not just compiled.

Still not covered by this manual check: `Open()` (only `Create` was exercised), error
paths (locked file, bad path, timeout), and everything beyond presentation lifecycle. The
"write a real integration test" step below is still the right next move — this was a
smoke check, not a test suite.

## What exists and builds today

- Solution skeleton: `Sbroenne.PowerPointMcp.slnx`, `Directory.Build.props`/`.targets`,
  `Directory.Packages.props` (central package management), `global.json`, `NuGet.Config`,
  `LICENSE` (MIT), `.gitignore`.
- **`src/PowerPointMcp.ComInterop`** — references `Microsoft.Office.Interop.PowerPoint`
  (real PIA, NuGet version `15.0.4420.1018` — the only version available on nuget.org at
  time of writing; re-check for newer releases before shipping), embedded via
  `ForceEmbedPowerPointInteropTypes` in `Directory.Build.targets` (verified: no
  `Microsoft.Office.Interop.PowerPoint.dll` or `office.dll` in build output).
  - `IOleMessageFilter.cs`, `OleMessageFilter.cs` — ported **verbatim** (namespace/doc-comment
    swap only) from mcp-server-excel. Fully Office-COM-generic, no Excel dependency existed
    in the original either.
  - `ComInteropConstants.cs`, `ComUtilities.cs` — ported, trimmed to the generic parts
    (timeouts, `Release`, `KernelSleep`, `TryQuitPowerPoint`). The Excel-specific `Find*`
    helpers (FindQuery/FindName/FindSheet/FindConnection) were **dropped** — PowerPoint
    equivalents (FindSlide, FindShape, ...) should be added per Core domain as needed.
  - `Session/PresentationContext.cs`, `IPresentationBatch.cs`, `PresentationBatch.cs`,
    `PresentationSession.cs`, `PresentationShutdownService.cs` — a **simplified,
    single-presentation port** of ExcelBatch/ExcelSession's STA-thread + channel-based
    work queue + OLE message filter + operation-timeout pattern. Builds and follows the
    same shape as the original, but has **not been tested against real PowerPoint COM**
    (no PowerPoint automation was run in this session — build-only verification).
- **`src/PowerPointMcp.Core`** — `Presentation/IPresentationCommands.cs` +
  `PresentationCommands.cs`, implementing only `Create()` and `Save()`, following the
  Success/ErrorMessage invariant and no-try-catch-suppression pattern from
  mcp-server-excel's Rule 1 / Rule 1b.
- **`src/PowerPointMcp.CLI`** — minimal hand-written `Program.cs` (not yet using
  Spectre.Console.Cli or the source-generator pattern) wiring `create <path>` end to end.
- `dotnet build` succeeds with 0 warnings/errors.

## Explicitly NOT done (do not assume otherwise)

- **No integration tests, no PowerPoint automation has actually been exercised.** Per this
  repo's Rule 30 equivalent (integration-tests-only, TDD), the very next step before
  writing more Core commands should be: write a real integration test that opens PowerPoint,
  creates a presentation, and verifies the file exists — run it and confirm PowerPoint
  actually launches correctly on this machine, THEN continue.
- No `PowerPointMcp.Service` (named-pipe daemon for CLI, in-process host for MCP Server).
- No `PowerPointMcp.McpServer` project at all yet.
- No `Generators`/`Generators.Cli`/`Generators.Mcp` — the CLI command above is hand-written,
  not generated.
- No Slide/Shape/TextFrame/Table/Chart/Image/Notes/Layout/Export-QA Core domains — only
  Presentation lifecycle (Create/Save; no Open/Close/List-sessions yet either).
- No multi-presentation batches, no IRM/AIP detection, no macro-security handling for
  `.pptm`, no resilient close/quit retry (`PresentationShutdownService` is a bare-bones
  inline close, not the exponential-backoff `ExcelShutdownService` equivalent).
- No `.github/instructions/*`, no pre-commit hook, no CI workflows, no skills docs.
- **The GitHub repo `sbroenne/mcp-server-powerpoint` does not exist yet** — this is a local
  git repo only (`D:\source\mcp-server-powerpoint`, branch `main`, no remote configured).
  The session's `gh` CLI was authenticated as an EMU account (`stbrnner_microsoft`) that
  cannot create repos under the personal `sbroenne` namespace — creating the remote repo
  needs to happen from an account with that access, then `git remote add origin ... && git push`.

## Recommended next steps (in order)

1. Create the real GitHub repo (`sbroenne/mcp-server-powerpoint`, MIT, public) and push
   this local history to it.
2. Write the first integration test (open real PowerPoint, create/save/verify a `.pptx`)
   and get it green — this validates the entire `PresentationBatch` STA/COM plumbing for
   real, which has not happened yet.
3. Fix whatever PowerPoint-COM quirks that first real test surfaces (expect some — this
   port was written by analogy to Excel, not verified against PowerPoint's actual COM
   behavior for `Presentations.Add`/`Open`/`SaveAs` parameter quirks).
4. Add `Open`/`Close`/session-list to `IPresentationCommands`, then start the next domain
   (Slide: add/delete/duplicate/reorder) following strict TDD.
5. Only after 2-3 domains are solid, invest in the CLI/MCP source generators (copy+adapt
   `ExcelMcp.Generators*` — do this once the Core interface shape has stabilized, not before).
6. Port `.github/instructions/*.instructions.md` and the pre-commit hook once there's real
   code for those gates to check.

See `plan.md` (session workspace) for the full architecture/domain-mapping plan and the
PowerPoint-MCP-server / Anthropic-pptx-skill research this project is based on.
