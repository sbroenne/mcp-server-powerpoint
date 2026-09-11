---
applyTo: "src/**/*.cs"
excludeAgent: "code-review"
---

# Architecture Patterns

> **Core patterns for PowerPointMcp development**

## .NET Class Design (MANDATORY)

1. **One Public Class Per File** - standard .NET practice.
2. **File Name = Class Name** - `ShapeCommands.cs` contains `ShapeCommands`.
3. **Partial Classes for Large Implementations** - split when a class grows past ~15 methods or
   spans multiple sub-concerns (not needed yet at current Core domain sizes).
4. **Descriptive Names** - no over-optimization (`ShapeCommands` not `Commands`).
5. **Folder = Organization, Not Identity** - `Commands`/domain folders group files; the type name
   itself doesn't encode the folder.

## Layered Architecture (CRITICAL)

```
ComInterop (STA thread, OLE message filter, PresentationBatch work queue)
    ↓
Core (domain commands: Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Master,
      Animation, Image, Media, Chart, Export — all but Presentation carry a [ServiceCategory] attribute)
    ↓
Service (PowerPointMcpService: session registry + dispatch, shared codebase)
    ↓                                              ↓
McpServer (in-process, ServiceBridge)        CLI / pptcli (named-pipe daemon, ServiceClient)
```

- **ComInterop** owns all direct COM object lifetime and STA-thread marshaling. Core never talks
  to `Microsoft.Office.Interop.PowerPoint` types outside of a `batch.Execute(...)` callback.
- **Core** commands take `IPresentationBatch batch` as an explicit first parameter (no ambient
  session state). Domains other than `Presentation` also carry a `[ServiceCategory("name",
  "PascalName")]` attribute (and an `[McpTool(...)]` attribute describing the generated tool) so
  `PowerPointMcp.Generators.Mcp`/`.Cli` can discover them — mirroring `mcp-server-excel`'s
  generator pipeline.
- **Service** (`PowerPointMcp.Service`) resolves `sessionId` → `IPresentationBatch` via its own
  `PresentationSessionRegistry`, then calls the matching Core command with the resolved batch.
  Both entry points call into this same dispatch logic (`ServiceRegistry.{Category}.RouteAction`)
  — MCP in-process via `ServiceBridge.ForwardToService`, CLI over a named pipe via
  `ServiceClient`/`IPowerPointDaemonRpc`.
- **McpServer** tools are thin either way: the single hand-written `presentation` tool
  (`PresentationTools.cs`) handles session/template/property operations via
  `PresentationSessionRegistry`; the generated action-dispatch tools
  (`slide`/`shape`/`chart`/etc., one per `[ServiceCategory]` domain) forward straight to the
  shared `PowerPointMcpService`. No domain logic lives in either.

## Command Pattern

### Structure (per domain)
```
Core/Shape/
├── IShapeCommands.cs     # Interface (contract) — [ServiceCategory] + [McpTool] attributes
├── ShapeCommands.cs      # Implementation
└── ShapeOperationResult.cs  # Result DTO (Success/ErrorMessage + domain fields)
```

### MCP Tool Routing (generated action-dispatch, one tool per domain)

Matching `mcp-server-excel`'s pattern, each `[ServiceCategory]` Core domain is exposed as a
**single generated action-dispatch MCP tool** (e.g. `shape`) taking an `action` parameter (e.g.
`"add-rectangle"`) rather than one hand-written `[McpServerTool]` method per verb. Do not
hand-write a new per-verb tool class for a Core domain — add the domain's
`[ServiceCategory]`/`[McpTool]` attributes and the generators emit the dispatch tool and its
`pptcli` commands automatically. `Presentation` remains hand-written in `PresentationTools.cs`,
but now follows the same single-tool action-dispatch shape.

```csharp
// Core: attribute-driven, discovered by the generators — no hand-written tool class needed.
[ServiceCategory("shape", "Shape")]
[McpTool("shape", Title = "Shape Operations", Destructive = true, Category = "content",
    Description = "Add, count, delete, reposition, and resize shapes on a slide in an open presentation session.")]
public interface IShapeCommands
{
    ShapeOperationResult AddRectangle(IPresentationBatch batch, int slideIndex, float left, float top, float width, float height);
    // ...
}
```

## Resource Management Pattern

All direct COM object access happens inside a `batch.Execute((ctx, ct) => { ... })` callback on
the `PresentationBatch`'s dedicated STA thread — never access
`Microsoft.Office.Interop.PowerPoint` types from arbitrary threads. Release any manually-obtained
COM references in a `finally` block, including intermediate typed COM objects.
NoPIA embedding removes the runtime PIA assembly dependency; it does not manage
COM reference lifetimes.

### PIA-First Interop (MANDATORY)

Follow the typed-access and late-binding requirements in
[critical rules](critical-rules.instructions.md). The
`ForceEmbedPowerPointInteropTypes` target embeds the PowerPoint and Office PIAs;
do not port Excel's runtime assembly resolver here.

If a typed Office enum is unavailable because only the PowerPoint PIA is embedded, prefer adding
the appropriate PIA/reference support over replacing the enum with an unexplained integer.

## Exception Propagation Pattern (CRITICAL)

**Core Commands: let exceptions propagate naturally** — do not suppress with a catch block that
returns an error result. See `critical-rules.instructions.md` Rule 1b for the
requirements, and the [development guide](../../docs/DEVELOPMENT-GUIDE.md)
for the rationale and validation example.

## Session Ownership (Service-Backed Registry, Two Entry Points)

`PowerPointMcp.Service`'s `PowerPointMcpService` owns a `PresentationSessionRegistry`
(`ConcurrentDictionary<string sessionId, IPresentationBatch>`) and is the single dispatch point
both entry points use — mirroring `mcp-server-excel`'s `ExcelMcp.Service` + `ServiceBridge`:

- **MCP Server** hosts `PowerPointMcpService` **in-process** (no pipe): DI singleton, called
  directly via `ServiceBridge.ForwardToService`. The stdio host is itself long-lived per client
  connection, so no RPC layer is needed on this side.
- **CLI** (`pptcli`) talks to a **separate background daemon process** hosting the same
  `PowerPointMcpService`, over a named pipe (`ServiceClient`/`IPowerPointDaemonRpc`,
  `StreamJsonRpc`), auto-started on first `session open`/`session create` and reused by every
  subsequent `pptcli` invocation referencing the same session id.

The two entry points run as **separate processes** with **separate PowerPoint instances** and do
**not** share live sessions with each other — only the same `Core`/`Service` codebase.

`presentation(action="create", filePath)` creates and saves the file and returns
an open session. Reuse that `sessionId`, rather than opening the file again.
`presentation(action="test", filePath)` validates that PowerPoint can open the
file without retaining a session. Presentation actions take `sessionId`;
generated domain tools take `session_id`.

For all-slide image export, prefer PowerPoint's single `Presentation.Export`
call over a loop of `Slide.Export` calls.
