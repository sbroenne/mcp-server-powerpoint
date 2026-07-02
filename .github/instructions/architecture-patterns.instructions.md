---
applyTo: "src/**/*.cs"
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
Core (domain commands: Presentation, Slide, Shape, TextFrame, Table, Notes, Layout, Image, Chart, Export)
    ↓
McpServer (hand-written [McpServerToolType] classes, PresentationSessionRegistry)
    ↓ (separately, not yet at parity)
CLI (placeholder Program.cs)
```

- **ComInterop** owns all direct COM object lifetime and STA-thread marshaling. Core never talks
  to `Microsoft.Office.Interop.PowerPoint` types outside of a `batch.Execute(...)` callback.
- **Core** commands take `IPresentationBatch batch` as an explicit first parameter (no ambient
  session state, no `[ServiceCategory]` marker attributes — those exist in `mcp-server-excel` for
  its generator pipeline, which this repo has NOT ported; see
  `.squad/decisions.md` for the rationale).
- **McpServer** tools are thin: resolve `sessionId` → `IPresentationBatch` via
  `PresentationSessionRegistry.TryGet`, call the matching Core command, serialize the
  `{Domain}OperationResult` to JSON. No domain logic lives in the `Tools/` classes.

## Command Pattern

### Structure (per domain)
```
Core/Shape/
├── IShapeCommands.cs     # Interface (contract)
├── ShapeCommands.cs      # Implementation
└── ShapeOperationResult.cs  # Result DTO (Success/ErrorMessage + domain fields)
```

### MCP Tool Routing (one static class per domain, not action-dispatch)

Unlike `mcp-server-excel`'s single action-dispatch tool per domain (e.g. `slide(action:'add')`),
PowerPointMcp registers **one `[McpServerTool]` method per verb** (e.g. `add_slide`,
`get_slide_count`, `delete_slide` as three separate MCP tools). This was a deliberate MVP choice
(see `.squad/decisions.md`, Q2) to keep the hand-written tool classes simple before generators are
introduced — do not silently convert existing tools to action-dispatch without an explicit
architecture decision.

```csharp
[McpServerToolType]
public static class ShapeTools
{
    private static readonly ShapeCommands Commands = new();

    [McpServerTool(Name = "add_rectangle")]
    [Description("Add a rectangle shape to the given slide.")]
    public static string AddRectangle(
        [Description("The session id returned by open_presentation.")] string sessionId,
        [Description("1-based index of the slide to add the rectangle to.")] int slideIndex,
        float left, float top, float width, float height,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_rectangle", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");

            return SerializeResult(Commands.AddRectangle(batch, slideIndex, left, top, width, height));
        });
}
```

## Resource Management Pattern

All direct COM object access happens inside a `batch.Execute((ctx, ct) => { ... })` callback on
the `PresentationBatch`'s dedicated STA thread — never access
`Microsoft.Office.Interop.PowerPoint` types from arbitrary threads. Release any manually-obtained
COM references in a `finally` block if the Core command holds intermediate `dynamic`/typed COM
objects beyond what the PIA's NoPIA embedding already manages for you.

## Exception Propagation Pattern (CRITICAL)

**Core Commands: let exceptions propagate naturally** — do not suppress with a catch block that
returns an error result. See `critical-rules.instructions.md` Rule 1b for the full rationale and
example.

## Session Ownership (In-Process Registry, No Out-of-Process Service)

Unlike `mcp-server-excel`'s `ExcelMcp.Service` + named-pipe `ServiceBridge`, PowerPointMcp's MVP
uses an **in-process singleton** `PresentationSessionRegistry`
(`ConcurrentDictionary<string sessionId, IPresentationBatch>`) registered as a DI singleton /
`IHostedService` inside the MCP stdio host itself. There is no separate daemon process, no named
pipe RPC — the STA thread + work queue already live inside `PresentationBatch`, and the MCP host
is itself long-lived per client connection.

Do not introduce an out-of-process Service without an explicit architecture decision — it's a
deliberate Phase-2 item (crash isolation, multi-client sharing, survives MCP restart), not
required for current functionality.
