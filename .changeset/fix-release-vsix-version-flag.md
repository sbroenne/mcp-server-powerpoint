---
"powerpointmcp": patch
---

Fix the release workflow's "Update Extension Version" step failing with "Version not changed"
whenever the release version coincidentally matches the VS Code extension's fixed baseline
version (0.1.0, never committed back to the repo since the bump is ephemeral per-run). Adds
`--allow-same-version` to the `npm version` call.
