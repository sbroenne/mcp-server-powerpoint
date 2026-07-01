# Squad Decisions

## Active Decisions

### 2026-07-01T11:07:03+02:00: Always use personal GitHub account (sbroenne)
**By:** sbroenne (via Copilot)
**What:** All git/gh operations in this repo must use the personal GitHub account `sbroenne`, NOT the EMU account `stbrnner_microsoft`. Repo creation, push, PRs, releases all go to the sbroenne namespace.
**Why:** User directive. The prior session was blocked from creating sbroenne/mcp-server-powerpoint because gh was authed as the EMU account which lacks sbroenne-namespace access.
**Status:** `gh auth status` confirms `sbroenne` is currently the ACTIVE account (keyring). The EMU account is signed in but inactive. Blocker is resolved. If gh ever shows stbrnner_microsoft as active, run `gh auth switch --user sbroenne` before any GitHub operation.

### 2026-07-01T11-14-34: PowerPointMcp.McpServer architecture — build order, session model, generators deferral

**By:** Dallas (Lead/Architect)
**What:** Architecture pass for building out the empty `src/PowerPointMcp.McpServer` project. Design only — no production code written.
**References:** src/PowerPointMcp.Core/*/I*Commands.cs, src/PowerPointMcp.ComInterop/Session/*, src/PowerPointMcp.McpServer (empty), src/PowerPointMcp.CLI/Program.cs, Directory.Packages.props; sibling D:\source\mcp-server-excel (ExcelMcp.McpServer, ExcelMcp.Service, ExcelMcp.Generators*).
**Why:** McpServer is the biggest gap in the port and the current focus. Grounding the design against real repo state + the proven Excel pattern prevents premature abstraction and keeps the port faithful where it should be.

SIBLING AVAILABILITY: mcp-server-excel IS present locally at D:\source\mcp-server-excel with a full MCP server, out-of-process Service, and source generators. Design below is anchored to its proven pattern, with PowerPoint deltas flagged.

KEY DECISIONS:

1. SDK / HOSTING — Port Excel's model exactly. `Host.CreateApplicationBuilder(args)` + `builder.Services.AddMcpServer(...).WithToolsFromAssembly()` + `.WithStdioServerTransport()`. Package `ModelContextProtocol` 0.5.0-preview.2 and `Microsoft.Extensions.Hosting` are ALREADY in Directory.Packages.props — no new packages required for the MVP. `net10.0-windows`, `OutputType=Exe`, `AssemblyName=Sbroenne.PowerPointMcp.McpServer`. Route ALL logging to stderr (`LogToStandardErrorThreshold = Trace`) — stdout is reserved for JSON-RPC. Skip Application Insights/telemetry for the MVP (defer).

2. SESSION OWNERSHIP — DIVERGE FROM EXCEL FOR THE MVP. Excel uses a separate out-of-process `ExcelMcp.Service` + `ServiceBridge` (named-pipe RPC); MCP tools forward `"category.action"` + sessionId strings. PowerPoint has NO Service project built. Decision: for the first cut, DO NOT build the out-of-process Service. Instead register an in-process singleton `PresentationSessionRegistry` (a `ConcurrentDictionary<string sessionId, IPresentationBatch>`) as a DI singleton / `IHostedService`. Rationale: the STA thread + channel queue already lives INSIDE `PresentationBatch`, and the MCP stdio host is itself long-lived per client — so an in-process registry fully satisfies "one long-lived session across many tool invocations" without RPC. The out-of-process Service (crash isolation, multi-client sharing, survives MCP restart) is a Phase-2 hardening port, not MVP.

3. SESSION LIFECYCLE / SHUTDOWN — New tool `presentation(action:'open'|'create')` calls `PresentationSession.BeginBatch/CreateNew`, stores the `IPresentationBatch` under a generated sessionId, returns the id. All other tools take `session_id`, look up the batch in the registry, and pass it to Core. `presentation(action:'close', save:bool)` saves (if requested), disposes the batch, removes it. CRITICAL: the host `finally`/`StopAsync` MUST dispose the registry → dispose every open batch → guarantees no lingering POWERPNT.exe on MCP client disconnect or Ctrl+C (mirrors Excel's `ServiceBridge.Dispose()` in Program.cs finally). Port `StdinPipeMonitor` so stdin EOF triggers graceful shutdown.

4. TOOLS: HAND-WRITE THE FIRST CUT, DEFER GENERATORS. The Excel generators discover Core interfaces via a `[ServiceCategory]` marker attribute and expect session-ID-based method signatures (Service resolves the batch). PowerPoint Core interfaces have NO such attributes and take `IPresentationBatch batch` as the first parameter directly. Standing up `ExcelMcp.Generators.Mcp` now would force a Core refactor (add attributes + swap batch→sessionId across 9 domains) before the MCP surface is even proven. Decision: Brett hand-writes ~9-10 static `[McpServerToolType]` tool classes (one per domain) that resolve session→batch from the registry and call Core. Only AFTER the tool surface is validated do we invest in porting `Generators.Shared/Common` + `Generators.Mcp` (Phase 2). This honors the charter: no generators until Core shape is proven stable across the MCP boundary.

5. RESULT SHAPE (Rule 1/1b → MCP). Core returns `{Domain}OperationResult` with `Success`/`ErrorMessage` (Rule 1: Success==true ⟹ ErrorMessage null). Tools serialize these to JSON via a shared `PowerPointToolsBase` (port of `ExcelToolsBase`): camelCase, `WriteIndented=false`, `IgnoreCondition=WhenWritingNull`, `JsonStringEnumConverter`, and add `isError=true` when `Success==false`. Expected bad input (bad index, missing file) already surfaces as `Success=false` from Core → returned as an error JSON payload, NOT thrown. Unexpected COM exceptions propagate and are caught ONLY at the tool boundary (`ExecuteToolAction` wrapper) which logs HResult to stderr and serializes a structured error (exceptionType, hresult, innerError). No try-catch suppression inside Core (Rule 1b preserved).

6. PROJECT REFERENCES / LAYOUT. McpServer.csproj references PowerPointMcp.Core (which transitively pulls ComInterop). NO Service reference (not built). NO generator reference (deferred). Port `Directory.Build.targets` office.dll/assembly-resolver handling so the PIA loads. Proposed layout:
   src/PowerPointMcp.McpServer/
     PowerPointMcp.McpServer.csproj
     Program.cs                 (host + stdio + shutdown)
     GlobalUsings.cs
     Session/PresentationSessionRegistry.cs   (in-process sessionId→IPresentationBatch)
     Tools/PowerPointToolsBase.cs             (JSON opts, ExecuteToolAction, error serialization)
     Tools/PresentationTool.cs                (open/create/save/close — owns session lifecycle)
     Tools/SlideTool.cs, ShapeTool.cs, TextFrameTool.cs, TableTool.cs,
     Tools/NotesTool.cs, LayoutTool.cs, ImageTool.cs, ChartTool.cs
     .mcp/server.json           (MCP manifest)

TASK BREAKDOWN FOR BRETT (ordered):
  T1. csproj + Program.cs stdio host (name "powerpoint-mcp"), stderr logging, ServerInstructions, graceful shutdown. Verify `--version`/`--help` and that it starts/stops with no lingering POWERPNT.exe.
  T2. PresentationSessionRegistry singleton + host-shutdown disposal of all batches.
  T3. PowerPointToolsBase (JSON options, ExecuteToolAction, SerializeToolError) — port from ExcelToolsBase minus telemetry.
  T4. PresentationTool (open/create/save/close) — proves the session lifecycle end-to-end (vertical slice).
  T5. SlideTool + ShapeTool — proves session→batch→Core call pattern for the multi-op domains.
  T6. Remaining domain tools: TextFrame, Table, Notes, Layout, Image, Chart.
  T7. .mcp/server.json manifest + README.
  (Phase 2, separate work: out-of-process Service + ServiceBridge; port Generators.* and replace hand-written tools.)

WHAT RIPLEY TESTS (strategy differs from Core's real-COM tests):
  - MCP-layer tests should use the SDK's in-memory pipe transport (Excel's `ConfigureTestTransport(inputPipe, outputPipe)` / `WithStreamServerTransport`) — port that hook into Program.cs so tests drive the server without stdio.
  - Protocol-level: tools list, schema, arg validation, and error-envelope shape (isError, no stdout pollution) can be asserted WITHOUT launching PowerPoint.
  - Session lifecycle: open→operate→close happy path + double-close/unknown-session error paths DO touch real COM (Rule 30) — keep serialized (xunit maxParallelThreads:1) and assert no lingering POWERPNT.exe after host shutdown.
  - Do NOT re-test Core command COM behavior at the MCP layer (already covered by Core integration tests); MCP tests assert wiring + serialization + session mapping only.

OPEN QUESTIONS FOR sbroenne (decide before/early in Brett's work):
  Q1. Confirm MVP skips the out-of-process Service (in-process registry only)? This is the main deliberate divergence from Excel.
  Q2. Tool granularity: one action-dispatch tool per domain (Excel style, e.g. `slide(action:...)`) vs one tool per verb? Recommend per-domain action dispatch to match Excel + keep the tool list small.
  Q3. "Show PowerPoint / Agent Mode" (BeginBatch already supports show:bool) in MVP or defer? Recommend defer to Phase 2.
  Q4. Telemetry / Application Insights in MVP? Recommend defer.
  Q5. Core currently lacks explicit `Open` and multi-presentation support (noted in interfaces) — confirm MVP is single-presentation-per-session, matching current Core scope.

### 2026-07-01T09-12-54: Recovered project roadmap after session crash + Alien-cast team hired
**By:** Squad-Coordinator
**What:** Recovered project roadmap after session crash + Alien-cast team hired
**References:** CONTINUATION.md, .squad/team.md, .squad/casting/registry.json
**Why:** Session crashed; session-workspace plan.md was lost. Roadmap reconstructed from CONTINUATION.md + git history and recorded here durably.

CURRENT STATE (2026-07-01):
- 9 Core domains done with real-COM TDD integration tests: Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart.
- ComInterop STA-batch plumbing validated against real PowerPoint; 2 real COM bugs found+fixed (empty-on-Add, Untitled=msoTrue breaking Save).
- PowerPointMcp.McpServer project folder exists but is EMPTY (zero .cs) — biggest gap.
- CLI is a hand-written placeholder Program.cs (not the target Generators-based CLI).
- No Service daemon, no Generators, no CI/hooks/instructions docs.
- Repo is LOCAL-ONLY: no GitHub remote. sbroenne/mcp-server-powerpoint does not exist yet. gh was authed as EMU account (stbrnner_microsoft) that cannot create repos under the sbroenne namespace.

NEXT STEPS (priority order):
1. Create GitHub remote (sbroenne/mcp-server-powerpoint, MIT, public) + push local history. BLOCKED on an account with sbroenne-namespace access.
2. Build the empty McpServer (Brett) — expose the 9 Core domains as MCP tools.
3. Add Open/Close/session-list + Export domain to Core (Parker).
4. Adapt ExcelMcp.Generators* now that Core shape is stable across 9 domains (Brett).
5. Port .github/instructions/*, pre-commit hook, CI (team).
6. Refresh CONTINUATION.md — its "NOT done" list is stale/misleading.

TEAM CAST (Alien universe): Dallas (Lead/Architect), Parker (COM/Core Dev), Brett (MCP/Service Dev), Ripley (Tester) + built-ins Scribe/Ralph/Rai.

CONVENTIONS: Rule 1/1b (Success/ErrorMessage invariant, no try-catch suppression). Rule 30 (real-COM integration tests only, no mocking, strict TDD). Tests serialized via xunit.runner.json maxParallelThreads:1.

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction

### 2026-07-01T11:30:00+02:00: Full release-deliverables surface in scope (port Excel's 8-job release.yml)
**By:** sbroenne (via Copilot)
**What:** Ship the COMPLETE distribution surface of mcp-server-excel — 10 deliverables:
1. CLI standalone single-file exe zip (pptcli.exe) — GitHub Release
2. CLI NuGet .NET tool (Sbroenne.PowerPointMcp.CLI) — NuGet.org
3. MCP Server standalone single-file exe zip (mcp-powerpoint.exe) — GitHub Release
4. MCP Server NuGet .NET tool (Sbroenne.PowerPointMcp.McpServer, PackageType=McpServer) — NuGet.org
5. VS Code extension VSIX (powerpoint-mcp, publisher sbroenne) contributing mcpServerDefinitionProviders + chatSkills — VS Code Marketplace (vsce)
6. MCPB bundle (.mcpb, Build-McpBundle.ps1) — Claude Desktop, GitHub Release
7. Agent Skills zip + manifest.json + CLAUDE.md + .cursorrules — npx skills add sbroenne/mcp-server-powerpoint
8. MCP Registry entry (.mcp/server.json via MCP Publisher) — official MCP Registry
9. GitHub Release with all artifacts + changelog + git tag v{ver}
10. Docs site (GitHub Pages) — Phase 4
**Version:** manual semver from latest tag; workflow_dispatch bump. **Secrets:** NUGET_USER (OIDC), VSCE_TOKEN, MCP Registry auth, GITHUB_TOKEN.
**Why:** User: "we also need vscode plugins etc" + "all the other release deliverables". Full distribution surface (which python-pptx rivals lack) is core to the "most complete" goal. Details: session plan.md §2b + §5a.

### 2026-07-01T11:32:40+02:00: GitHub Pages docs site in scope — domain powerpointmcpserver.dev
**By:** sbroenne (via Copilot)
**What:** Ship a GitHub Pages docs site (deliverable #10). Port mcp-server-excel/gh-pages/ (Jekyll, Ruby 3.3, build.sh production) + deploy-gh-pages.yml (deploys on push to main touching gh-pages/**, FEATURES.md, CHANGELOG.md; env github-pages; pages:write + id-token:write). IndexNow SEO (Bing/Yandex) with new key file. Content: landing page, feature matrix, 6-path install matrix, quickstart, tool/command reference, changelog.
**Custom domain:** CONFIRMED = **powerpointmcpserver.dev** (sbroenne, 2026-07-01). Needs DNS (A/ALIAS to GitHub Pages IPs or CNAME to sbroenne.github.io) + CNAME file in gh-pages/ + Enforce HTTPS. Landing page: https://powerpointmcpserver.dev
**Why:** User: "we also need github pages" + "the domain will be powerpointmcpserver.dev". Details: session plan.md Phase 4 + §5a row 10.
