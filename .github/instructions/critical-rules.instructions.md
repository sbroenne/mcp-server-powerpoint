---
applyTo: "**"
excludeAgent: "code-review"
---

# Critical rules

## Rule 1: Success and errors

`Success == true` requires `ErrorMessage == null`. Set success only on the
success path. Expected caller-correctable failures (bad indices, missing files,
unknown sessions) return `Success=false`; validate before the failing COM call.

## Rule 1b: Exception propagation

Never catch exceptions from `batch.Execute()` in Core and turn them into error
results. Unexpected COM/runtime exceptions propagate through the batch.
`PowerPointToolsBase.ExecuteToolAction` is the MCP catch-all boundary: it logs
the HResult to stderr and returns a structured error. Do not add another
catch-all in individual tools. MCP stdout is reserved for JSON-RPC.

## COM access and ownership

- All PowerPoint access runs inside `batch.Execute` on its STA thread.
- Use typed PowerPoint/Office PIA members and enums. Before using `dynamic`,
  reflection, raw dispatch, or numeric enum values, confirm that the restored
  interop metadata lacks the typed API, document the missing surface, and cover
  the exception with a real-COM test. Convenience is not an exception.
- Release every manually acquired COM reference in `finally`, including typed
  references. Embedded PIA types do not manage COM reference lifetimes.
- Slide, shape, row, and column indices are 1-based. Zero and negative values
  are validation failures, not alternative indexing.

## Rule 30: Behavior changes need real evidence

Write a focused failing regression test, observe the expected failure, then
implement and rerun it. Never mock COM for Core commands: use real PowerPoint.
Keep COM tests serialized and use explicit execution timeouts.
Protocol-only MCP tests may use the SDK in-memory transport.
See [testing strategy](testing-strategy.instructions.md) for filters and scope.

## Session lifecycle

- Track every live session in `PresentationSessionRegistry`. On host shutdown,
  both `PresentationSessionShutdownService.StopAsync` and `Main`'s `finally`
  call the idempotent `DisposeAll()` backstop.
- Close removes a session immediately and disposes its batch in the background;
  do not block close on process exit. PowerPoint may take minutes to exit after
  Quit. Do not force-kill on the happy path.
- Cleanup targets only proven process ownership (PID plus start time), never
  process names, window-title matches, or another user's PowerPoint.

## Data and authorization

Keep customer names, presentation contents, credentials, and private paths out
of commits, PRs, issues, logs, and other public artifacts. Use generic examples.
Keep temporary notes outside the repository. Never commit, push, merge, or
publish without explicit user authorization; see the root Git guidance for
GitHub PR assignments.
