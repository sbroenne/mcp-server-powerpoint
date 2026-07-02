---
applyTo: "tests/**/*.cs"
---

# Testing Strategy

> **Rule 30: real-COM integration tests only.** There are no unit tests with mocked COM objects
> anywhere in this repo's Core test suite — every test that touches `Core` commands drives a real
> PowerPoint desktop instance.

## Strict TDD (Red → Green)

1. Write a failing test that proves the bug/gap exists (red).
2. Run it, confirm it fails for the expected reason (not a compile error or unrelated failure).
3. Implement the minimal Core/ComInterop change to make it pass (green).
4. Re-run the specific test, then the domain's full `Feature=` filter, to confirm no regression.

Do not write implementation before a failing test exists for it.

## Test Project Layout

```
tests/
├── PowerPointMcp.Core.Tests/         # Real-COM integration tests, one file per domain
│   ├── xunit.runner.json             # maxParallelThreads: 1
│   ├── ShapeCommandsTests.cs          [Trait("Category","Integration")] [Trait("Feature","Shape")]
│   ├── ChartCommandsTests.cs          [Trait("Feature","Chart")]
│   └── ...
└── PowerPointMcp.McpServer.Tests/     # MCP protocol + session round-trip tests
    ├── xunit.runner.json
    └── Integration/
        ├── McpProtocolTests.cs        # No COM — in-memory transport, schema/wiring assertions
        └── McpRoundTripTests.cs       # Real COM — full open→edit→save→close lifecycle
```

## Serialization Is Mandatory

`xunit.runner.json` sets `maxParallelThreads: 1` in both test projects. PowerPoint does not
support multiple concurrent automation sessions launched by parallel test threads reliably —
running tests in parallel WILL cause flaky failures or hangs. Never remove or override this
setting to "speed up" a local run.

**Anti-pattern:** using both `IClassFixture<T>` and `[Collection("...")]` with a collection
fixture on the same test class — dual fixtures can create concurrent Excel/PowerPoint sessions
that deadlock during initialization under `maxParallelThreads: 1`. Use only the collection fixture
if a shared fixture is needed.

## Trait Tagging (Required for Surgical Filtering)

Every Core integration test class MUST carry both traits:

```csharp
[Trait("Category", "Integration")]
[Trait("Feature", "Shape")]
public class ShapeCommandsTests
{
    // ...
}
```

`Feature` values match the Core domain folder name (`Presentation`, `Slide`, `Shape`, `TextFrame`,
`Table`, `Notes`, `Layout`, `Image`, `Chart`, `Export`). This enables:

```powershell
# Fast, targeted feedback for the domain you're changing
dotnet test tests\PowerPointMcp.Core.Tests --filter "Feature=Shape"

# Full Core suite (real COM, serialized — slower)
dotnet test tests\PowerPointMcp.Core.Tests

# MCP protocol + round-trip suite
dotnet test tests\PowerPointMcp.McpServer.Tests
```

Always run tests with an explicit timeout in the terminal/tooling layer that invokes `dotnet
test` — never leave a COM test run open-ended; fail fast if PowerPoint stalls instead of hanging
the whole session.

## MCP-Layer Testing Strategy (Differs From Core)

- **Protocol-level tests** (`McpProtocolTests`) use the MCP SDK's in-memory pipe transport
  (`Program.ConfigureTestTransport` / `WithStreamServerTransport`) to drive the server WITHOUT
  launching PowerPoint. Use these for: `tools/list` shape, JSON schema correctness (e.g.
  confirming the DI-injected `registry` parameter never leaks into the schema), and error-envelope
  shape (`isError`, no stdout pollution from logging).
- **Session lifecycle tests** (`McpRoundTripTests`) DO touch real COM: open → operate → close
  happy path, plus double-close/unknown-session error paths. Keep these serialized like Core
  tests.
- **Do NOT re-test Core command COM behavior at the MCP layer** — that's already covered by Core
  integration tests. MCP tests assert wiring + serialization + session mapping only.

## Benign Office Shutdown Latency (Not a Bug)

After a session's `close_presentation`/batch disposal, the underlying `POWERPNT.exe` process can
take anywhere from ~90 to 200+ seconds to actually exit the OS process list — this is documented
Office post-Quit cleanup/telemetry behavior, not a leak. It self-resolves with no manual
intervention and does not block back-to-back test runs. Do NOT add a force-kill on the happy path
in test teardown; a force-kill safety net (if one exists) is reserved for genuine
`_operationTimedOut` cases only.

## Diagnose Before Coding (Golden Rule)

No production code changes without a failing test first that proves the bug exists. Diagnose the
root cause (which layer: ComInterop STA/thread issue vs. Core command logic vs. MCP tool
serialization) before writing a fix — a fix at the wrong layer wastes an entire session and gets
reverted.

## Bug Fix Pattern Search

When fixing a bug in one Core command, search for the same pattern across sibling commands in
other domains (e.g., an off-by-one in `Shape`'s index validation might also exist in `Table` or
`TextFrame`). Fix all matching cases in the same PR, or explicitly document why a similar-looking
case is not affected.
