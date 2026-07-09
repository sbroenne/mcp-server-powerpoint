---
applyTo: "src/PowerPointMcp.McpServer/**/*.cs"
---

# MCP Server Development Guide

> All tool methods in this project are **synchronous** — they return `string`, not
> `Task<string>`. There is no async/await anywhere in the `Tools/` classes; COM access via
> `IPresentationBatch.Execute` is already synchronous from the caller's point of view (the async
> plumbing lives inside `PresentationBatch`, not in the tool methods).

## LLM-Facing Content Rules

**NO EMOJIS in LLM-consumed content** — never use emoji characters in:
- Tool `[Description(...)]` attributes and XML `/// <summary>` comments — the MCP SDK extracts
  these into the tool schema an LLM reads directly.
- `skills/shared/*.md` and any future MCP prompt content.

**Use plain text markers:** "IMPORTANT:", "WARNING:", "NOTE:", "CRITICAL:".

**DO keep emojis in user-facing content:** README.md, this repo's own governance docs, and other
human-read documentation — humans appreciate visual aids; LLM tool schemas do not need them and
they cost tokens for no benefit.

## Two Kinds of Tools: Hand-Written vs. Generated

Most of the MCP tool surface is **generated**, not hand-written. Before editing anything under
`Tools/`, know which kind you're touching:

- **Hand-written** (`PresentationTools.cs` only): `create_presentation`, `open_presentation`,
  `save_presentation`, `close_presentation`, `list_sessions`, `apply_template`, `get_theme_name` —
  session lifecycle and template application don't fit the per-session action-dispatch shape, so
  they stay hand-written following the pattern below.
- **Generated** (everything else — `slide`, `shape`, `textframe`, `table`, `notes`, `layout`,
  `master`, `animation`, `image`, `chart`, `export`): one action-dispatch tool per
  `[ServiceCategory]` Core domain, emitted by `PowerPointMcp.Generators.Mcp` from the Core
  interface's `[ServiceCategory]`/`[McpTool]` attributes and XML doc comments. **Never hand-write a
  new tool class for one of these domains** — add the operation to the Core interface (with XML
  docs) and the generator picks it up. See `architecture-patterns.instructions.md`'s Command
  Pattern section for the Core-side attribute shape.

The rest of this guide (verb-per-tool pattern, manual `SerializeResult`, etc.) applies to the
hand-written `PresentationTools.cs` tools only.

## Implementation Pattern: One Verb Per Tool (Hand-Written Tools Only)

```csharp
[McpServerToolType]
public static class PresentationTools
{
    private static readonly PresentationCommands Commands = new();

    [McpServerTool(Name = "create_presentation")]
    [Description("Create a new, blank presentation and save it to disk.")]
    public static string CreatePresentation(
        [Description("Path to save the new presentation to.")] string filePath)
        => PowerPointToolsBase.ExecuteToolAction("create_presentation", () =>
        {
            return SerializeResult(Commands.Create(filePath));
        });

    private static string SerializeResult(PresentationOperationResult result) => /* ... */;
}
```

Every hand-written tool class follows this shape: a
`private static readonly {Domain}Commands Commands = new();` field, one `[McpServerTool]` method
per verb wrapped in `PowerPointToolsBase.ExecuteToolAction`, and a shared `SerializeResult` helper
for the domain's `{Domain}OperationResult` shape. Keep new hand-written tools consistent with this
pattern rather than inventing a new one — but remember: this only applies to
`PresentationTools.cs`. A new Shape/Chart/etc. operation goes into Core, not here.

## Error Handling (MANDATORY)

**MCP tools must return JSON with `isError: true` for business errors, NOT throw exceptions.**
This follows the MCP spec's two error mechanisms:

1. **Protocol errors** — malformed requests, unknown tools → the MCP SDK handles these; tool code
   rarely needs to throw for this.
2. **Tool execution errors** (business logic failures — unknown session, bad index, missing file)
   → return a JSON payload with `isError: true` via `PowerPointToolsBase.ValidationError(...)`, do
   NOT throw.

```csharp
// CORRECT — expected bad input: return a validation error payload
if (!registry.TryGet(sessionId, out var batch))
{
    return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");
}

// CORRECT — Core Success=false result: serialize as-is (already has errorMessage + isError)
return SerializeResult(Commands.Delete(batch, slideIndex));

// WRONG — throwing for an expected, caller-correctable condition
if (!registry.TryGet(sessionId, out var batch))
{
    throw new InvalidOperationException($"Unknown sessionId: {sessionId}");
}
```

**Unexpected exceptions** (COM exceptions, null refs, etc.) are allowed to propagate out of the
tool method body — `PowerPointToolsBase.ExecuteToolAction` wraps every tool call and catches them
at that single boundary, logging the HResult to stderr and serializing a structured error via
`SerializeToolError`. Do not add a second try-catch inside an individual tool method — let
`ExecuteToolAction` be the only catch-all.

## Session Injection Pattern (Hand-Written Tools)

`PresentationSessionRegistry registry` is a **plain parameter**, not a tool-facing argument — the
MCP SDK (1.3.0+) resolves it from the DI container and correctly excludes it from the generated
JSON schema (verified via `tools/list`). Every hand-written tool that needs to look up a session
takes it as the last parameter, named exactly `registry`. Generated action-dispatch tools instead
take a DI-injected `PowerPointMcpService service` parameter, which the generator wires
automatically — you never write this by hand.

## Result Serialization Pattern

Each domain tool class has a private `SerializeResult({Domain}OperationResult result)` helper
that projects the Core result DTO into the MCP JSON payload:

```csharp
private static string SerializeResult(ShapeOperationResult result)
    => PowerPointToolsBase.Serialize(new
    {
        success = result.Success,
        errorMessage = result.ErrorMessage,
        shapeIndex = result.ShapeIndex,
        shapeCount = result.ShapeCount,
        isError = result.Success ? (bool?)null : true
    });
```

`PowerPointToolsBase.JsonOptions` already applies camelCase naming, omits null properties
(`DefaultIgnoreCondition.WhenWritingNull`), and serializes enums as strings — don't duplicate that
configuration per tool class.

## Adding a New Tool

**For a generated domain (Slide, Shape, TextFrame, Table, Notes, Layout, Master, Animation,
Image, Chart, Export) — the common case:**
1. Add the Core command + `{Domain}OperationResult` fields first (Core-first, tested with real
   COM per `testing-strategy.instructions.md`), with an XML doc `<summary>` — the generator uses
   it as the operation's description.
2. Nothing else to write by hand — `PowerPointMcp.Generators.Mcp` picks up the new interface
   method automatically and adds it as a new `operation` value on that domain's action-dispatch
   tool (e.g. `shape.add-oval`) the next time the project builds.
3. Verify the new operation appears correctly in `tools/list`'s schema (protocol test in
   `tests/PowerPointMcp.McpServer.Tests`) and that `PowerPointMcp.Generators.Cli` emitted the
   matching `pptcli {category} {action}` command.
4. Update `skills/shared/*.md` (and its copy under `skills/powerpoint-mcp/references/`) if the new
   operation changes recommended workflows.

**For a hand-written tool (`PresentationTools.cs` only) — rare, session-lifecycle/template work:**
1. Add the Core command first, same as above.
2. Add the `[McpServerTool(Name = "snake_case_name")]` method to `PresentationTools.cs`, following
   the pattern above.
3. Verify the new tool appears correctly in `tools/list` with the expected parameter schema and no
   leaked `registry` parameter.
4. Update `skills/shared/*.md` as above.
