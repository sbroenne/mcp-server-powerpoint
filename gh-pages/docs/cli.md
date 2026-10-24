---
title: CLI Reference
description: Full CLI command reference for pptcli, the token-efficient PowerPoint automation CLI for coding agents.
---

# CLI Reference

## Interrupted Commands and Service Shutdown

Cancelling a command or reaching its wait limit does not close an existing session
or undo changes. PowerPoint may still be working. Run `pptcli session list` before
retrying; `operation.executionPending` shows unfinished work and `operation.errorMessage`
reports a later failure. Status and session lists remain available while an edit is busy.

`pptcli service stop` leaves unsaved or unfinished presentations open without saving
them. `--force` stops only the verified service process, never PowerPoint.
Explicit `pptcli session close <id>` still discards changes unless `--save` is supplied.

--8<-- "_generated/cli.md"
