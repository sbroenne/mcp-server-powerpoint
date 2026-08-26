---
applyTo: "**"
excludeAgent: "cloud-agent"
---

# Copilot Code Review

Report only actionable, high-confidence defects introduced by the pull request. Prioritize
correctness, data loss, resource leaks, deadlocks, broken contracts, and security. Do not request
style-only changes, speculative refactors, or unrelated cleanup.

## Interop and Resource Safety

- Enforce PIA-first access across Core and ComInterop. Flag new `dynamic` PowerPoint access when
  `Microsoft.Office.Interop.PowerPoint` exposes a typed member. Allow late binding only when a
  concise comment identifies the missing PIA surface and a real-PowerPoint test covers it.
- Verify every acquired COM object is released in a `finally` block. Cleanup must target tracked
  PID-plus-start-time identities, never a process name or window-title substring.
- Do not treat a generic HRESULT as proof of one specific cause without another distinguishing
  check. Do not add catch-all blocks inside Core commands; unexpected failures must propagate
  through `batch.Execute` to the MCP boundary.
- Verify counters, flags, locks, and session state are restored on every exit path. A timeout or
  connection failure must not become an empty successful result or strand PowerPoint.

## Contract and Path Completeness

- Trace action and parameter changes through the Core interface, generated service dispatch, CLI
  routing, MCP schema, help, skills, and both entry points. Names, types, aliases, defaults,
  validation, and timeout behavior must agree. Unknown values must fail rather than select a
  default.
- When a change adds a dependency, response field, wait, lock, or synchronization path, verify
  the corresponding existence guard, integrity check, return-value check, timeout, and round-trip
  assertion.
- If one generated, templated, or intentionally parallel artifact changes, verify its source of
  truth and counterparts change consistently. Flag copied logic and checks that compare generated
  output against the same stale fallback used to produce it.

## Protocol, Tests, and User-Facing Surfaces

- Reserve MCP stdout exclusively for JSON-RPC. Reachable logs, diagnostics, installers, and
  bootstrap messages must use stderr.
- After renames or behavior changes, search for stale names, flags, defaults, summaries, help,
  logs, skills, and documentation. Counts and versions must come from authoritative metadata or
  be guarded mechanically.
- Core tests must use real PowerPoint, use 1-based indices, and assert the resulting presentation
  state and returned fields rather than only `Success`.
- A success result must never include an error message. Expected input failures return
  `Success=false`; unexpected COM/runtime failures propagate.
