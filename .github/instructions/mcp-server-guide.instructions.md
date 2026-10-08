---
applyTo: "src/PowerPointMcp.McpServer/**/*.cs,src/PowerPointMcp.Generators.Mcp/**/*.cs,src/PowerPointMcp.Generators.Shared/**/*.cs"
excludeAgent: "code-review"
---

# MCP Server Development Guide

Tool methods return the MCP SDK's `Task<CallToolResult>`. Results include both a
text JSON block for existing clients and the same object as structured content.
Generated tools declare an output schema derived from their Core result
contracts. A cancellation token is injected by the SDK and forwarded through
the service bridge; it is never advertised as a tool argument.

## LLM-Facing Content Rules

**NO EMOJIS in LLM-consumed content** — never use emoji characters in:
- Tool `[Description(...)]` attributes and XML `/// <summary>` comments — the MCP SDK extracts
  these into the tool schema an LLM reads directly.
- skill files and any future MCP prompt content.

**Use plain text markers:** "IMPORTANT:", "WARNING:", "NOTE:", "CRITICAL:".

**DO keep emojis in user-facing content:** README.md, this repo's own governance docs, and other
human-read documentation — humans appreciate visual aids; LLM tool schemas do not need them and
they cost tokens for no benefit.

## Two Kinds of Tools: Hand-Written vs. Generated

Most of the MCP tool surface is **generated**, not hand-written. Before editing anything under
`Tools/`, know which kind you're touching:

- **Hand-written** (`PresentationTools.cs` only): the single `presentation` MCP tool. It is
  action-dispatch like Excel's file tool, but stays hand-written because create/open/list/close
  need custom session-registry behavior and optional `presentation_session_id`.
- **Generated** (everything else — `slide`, `shape`, `textframe`, `table`, `notes`, `layout`,
  `master`, `animation`, `image`, `media`, `chart`, `smartart`, `export`, `pagesetup`,
  `accessibility`, `customshow`): one action-dispatch tool per
  `[ServiceCategory]` Core domain, emitted by `PowerPointMcp.Generators.Mcp` from the Core
  interface's `[ServiceCategory]`/`[McpTool]` attributes and XML doc comments. **Never hand-write a
  new tool class for one of these domains** — add the operation to the Core interface (with XML
  docs) and the generator picks it up. Interfaces that declare `[McpReadOnlyActions(...)]` also
  get a `{tool}_read` MCP alias containing only those actions. The original tool remains intact,
  and the CLI keeps its existing command surface. See `architecture-patterns.instructions.md`'s
  Command Pattern section for the Core-side attribute shape.

The rest of this guide applies to the hand-written `presentation` tool only.

## Implementation Pattern: Single Hand-Written Dispatch Tool

```csharp
[McpServerToolType]
public static class PresentationTools
{
    private static readonly PresentationCommands Commands = new();

    [McpServerTool(Name = "presentation")]
    [Description("Presentation lifecycle, template, and document-property operations.")]
    public static Task<CallToolResult> Presentation(
        PresentationToolAction action,
        string? filePath = null,
        string? presentation_session_id = null,
        PresentationSessionRegistry? registry = null,
        CancellationToken cancellationToken = default)
        => PowerPointToolsBase.ExecuteToolActionAsync(
            "presentation", action.ToActionString(), () =>
        {
            return action switch
            {
                PresentationToolAction.Create => HandleCreate(filePath, false, registry!),
                PresentationToolAction.Open => HandleOpen(filePath, registry!),
                _ => PowerPointToolsBase.ValidationError($"Unknown action: {action}")
            };
        }, cancellationToken);
}
```

`PresentationTools.cs` is a single hand-written dispatch tool, not one method-per-verb. Keep new
presentation-lifecycle/template/property work inside that switch. A new Shape/Chart/etc.
operation still goes into Core, not here.

## Error Handling (MANDATORY)

**MCP tools must return JSON with `isError: true` for business errors, NOT throw exceptions.**
This follows the MCP spec's two error mechanisms:

1. **Protocol errors** — malformed requests and unknown tools are handled before dispatch by the
   MCP SDK and `ToolArgumentFilter`. The filter returns an MCP error result for unknown actions,
   unknown parameters, primitive type mismatches, and parameters invalid for the selected action.
2. **Tool execution errors** (business logic failures — unknown session, bad index, missing file)
   → return a JSON payload with `isError: true` via `PowerPointToolsBase.ValidationError(...)`, do
   NOT throw.

```csharp
// CORRECT — expected bad input: return a validation error payload
if (!registry.TryGet(presentation_session_id, out var batch))
{
    return PowerPointToolsBase.ValidationError($"Unknown presentation_session_id: {presentation_session_id}");
}

// CORRECT — Core Success=false result: serialize as-is (already has errorMessage + isError)
return SerializeResult(Commands.Delete(batch, slideIndex));

// WRONG — throwing for an expected, caller-correctable condition
if (!registry.TryGet(presentation_session_id, out var batch))
{
    throw new InvalidOperationException($"Unknown presentation_session_id: {presentation_session_id}");
}
```

**Unexpected exceptions** (COM exceptions, null refs, etc.) are allowed to propagate out of the
tool method body — `PowerPointToolsBase.ExecuteToolActionAsync` wraps every tool call and catches them
at that single boundary, logging the HResult to stderr and serializing a structured error via
`SerializeToolError`. Do not add a second try-catch inside an individual tool method — let
`ExecuteToolActionAsync` be the only catch-all. Cancellation is rethrown to the SDK rather than
reported as a PowerPoint failure.

## Session Injection Pattern (Hand-Written Tools)

`PresentationSessionRegistry registry` is a **plain parameter**, not a tool-facing argument — the
MCP SDK (1.3.0+) resolves it from the DI container and correctly excludes it from the generated
JSON schema (verified via `tools/list`). Every hand-written tool that needs to look up a session
takes it as the last parameter, named exactly `registry`. Generated action-dispatch tools instead
take a DI-injected `PowerPointMcpService service` parameter, which the generator wires
automatically — you never write this by hand.

## Result and Schema Pattern

Each generated domain tool projects the Core result DTO into compact JSON, then
`PowerPointToolsBase.CreateToolResult` returns it as both legacy text and structured content:

```csharp
private static string SerializeResult(ShapeOperationResult result) =>
    PowerPointToolsBase.Serialize(new
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
configuration per tool class. The MCP generator also emits one output-schema class per tool from
the public Core result properties. Keep result contracts accurate rather than hand-maintaining
generated schema.

`UseStructuredContent = true` and `OutputSchemaType` are mandatory. Preserve the text content for
older clients; structured content is additive.

## Adding a New Tool

**For a generated domain (Slide, Shape, TextFrame, Table, Notes, Layout, Master, Animation,
Image, Media, Chart, Export, CustomShow) — the common case:**
1. Add the Core command + `{Domain}OperationResult` fields first (Core-first, tested with real
   COM per `testing-strategy.instructions.md`), with an XML doc `<summary>` — the generator uses
   it as the operation's description.
2. Nothing else to write by hand — `PowerPointMcp.Generators.Mcp` picks up the new interface
   method automatically and adds it as a new `action` value on that domain's action-dispatch
   tool (e.g. `shape(action: "add-oval", ...)`) the next time the project builds.
3. Verify the new operation appears correctly in `tools/list`, its output fields come from the
   Core result contract, malformed action arguments are rejected before dispatch, and
   `PowerPointMcp.Generators.Cli` emitted the matching `pptcli {category} {action}` command.
4. Update the relevant page under `gh-pages/docs/reference/` if the new operation changes
   recommended workflows. Keep the compact skill focused on tool discovery and safe-use basics.

**For a hand-written tool (`PresentationTools.cs` only) — rare, session-lifecycle/template work:**
1. Add the Core command first, same as above.
2. Add the new enum value + switch arm to `PresentationTools.cs`, following the pattern above.
3. Update `PresentationToolOutputSchema` when the action adds a new result field, then verify the
   `presentation` tool appears in `tools/list` with structured output and no leaked `registry` or
   `cancellationToken` parameter.
4. Update the relevant documentation-site page as above.
