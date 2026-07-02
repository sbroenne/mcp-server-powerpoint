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

## Implementation Pattern: One Verb Per Tool

```csharp
[McpServerToolType]
public static class SlideTools
{
    private static readonly SlideCommands Commands = new();

    [McpServerTool(Name = "add_slide")]
    [Description("Add a new blank slide at the end of the presentation.")]
    public static string AddSlide(
        [Description("The session id returned by open_presentation.")] string sessionId,
        PresentationSessionRegistry registry)
        => PowerPointToolsBase.ExecuteToolAction("add_slide", () =>
        {
            if (!registry.TryGet(sessionId, out var batch))
                return PowerPointToolsBase.ValidationError($"Unknown sessionId: {sessionId}");

            return SerializeResult(Commands.AddBlank(batch));
        });

    private static string SerializeResult(SlideOperationResult result) => /* ... */;
}
```

Every tool class follows this shape: a `private static readonly {Domain}Commands Commands = new();`
field, one `[McpServerTool]` method per verb wrapped in `PowerPointToolsBase.ExecuteToolAction`,
and a shared `SerializeResult` helper for the domain's `{Domain}OperationResult` shape. Keep new
domain tool classes consistent with this pattern rather than inventing a new one.

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

## Session Injection Pattern

`PresentationSessionRegistry registry` is a **plain parameter**, not a tool-facing argument — the
MCP SDK (1.3.0+) resolves it from the DI container and correctly excludes it from the generated
JSON schema (verified via `tools/list`). Every tool that needs to look up a session takes it as
the last parameter, named exactly `registry`.

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

1. Add the Core command + `{Domain}OperationResult` fields first (Core-first, tested with real
   COM per `testing-strategy.instructions.md`).
2. Add the `[McpServerTool(Name = "snake_case_name")]` method to the matching `Tools/*.cs` class,
   following the pattern above.
3. Verify the new tool appears correctly in `tools/list` (protocol test in
   `tests/PowerPointMcp.McpServer.Tests`) with the expected parameter schema and no leaked
   `registry` parameter.
4. Update `skills/shared/*.md` (and its copy under `skills/powerpoint-mcp/references/`) if the new
   tool changes recommended workflows.
