# Release Strategy

PowerPointMcp releases all components together (MCP Server, CLI, npm launchers
and Windows runtimes, VS Code Extension, MCPB, Agent Skills, MCP Registry entry)
under a single version number, via the
`.github/workflows/release.yml` `workflow_dispatch` workflow.

## Cutting a release

1. Trigger `Release All Components` from the Actions tab with a `version_bump`
   (major/minor/patch) or a `custom_version`.
2. The workflow calculates the next version from the latest git tag and generates
   tool, operation, domain, and per-domain operation counts from the built code.
3. It applies those generated documentation updates to every package, builds and
   publishes every component (npm, NuGet, standalone exe zips, VS Code
   Marketplace, MCPB, Agent Skills zip, MCP Registry), creates the git tag,
   then creates the GitHub Release. The npm Windows runtime packages are
   published before their architecture-selecting launcher packages.

## Changelog generation

`CHANGELOG.md` is generated from **changesets** (`@changesets/cli`), not hand-edited.

- Contributors add a small markdown fragment under `.changeset/` describing their
  user-facing change (`npx changeset` — see [`.changeset/README.md`](../.changeset/README.md)).
  A CI check (`.github/workflows/changeset-check.yml`) fails a PR that changes
  user-facing behavior but has no changeset and no `skip-changelog` label.
- At release time, the `create-release` job in `release.yml` runs
  [`scripts/Build-Changelog.ps1`](../scripts/Build-Changelog.ps1), which:
  1. Runs `npx changeset version` to consume all pending `.changeset/*.md`
     fragments into a new section at the top of `CHANGELOG.md`.
  2. Normalizes that section's header to this repo's `## [X.Y.Z] - YYYY-MM-DD`
     (Keep a Changelog) style.
  3. Forces the bookkeeping-only root `package.json` version to match the real
     release version (the source of truth is the git tag / workflow input, not
     `package.json`).
  4. Extracts the new section's body to `release_notes_body.md`, which is
     substituted into [`.github/release-notes-template.md`](../.github/release-notes-template.md)
     to produce the GitHub Release body.
- A follow-up step opens a PR (`chore/changelog-v<version>`) committing the
  updated documentation counts, `CHANGELOG.md`, `package.json`, and consumed
  `.changeset/*.md` deletions back to `main`, since branch protection prevents a direct push.
  This step deliberately does **not** use `continue-on-error` — if it fails
  (e.g. missing permissions), the release is left visibly incomplete rather
  than silently missing its changelog commit-back.

### Why changesets instead of hand-edited `CHANGELOG.md`

The previous approach relied on contributors manually editing a `## [Unreleased]`
section, then an `awk`/`sed`-based extraction step in `release.yml` to pull that
section into the GitHub Release body. This was fragile: entries were easy to
forget, mis-format, or leave permanently mislabeled as `[Unreleased]` if the
extraction step silently failed. Changesets (used by React, Remix, Vite, and many
other open source projects) make the changelog entry part of the PR itself,
enforced by CI, and compiled deterministically at release time.

### Node/npm in a .NET repo

The root `package.json` and `.changeset/` host `@changesets/cli` for
`CHANGELOG.md` generation. The `npm-packages/` tree additionally contains the
shared Node launcher and package metadata for:

- `@sbroenne/mcp-server-powerpoint`
- `@sbroenne/pptcli`
- their Windows x64 and ARM64 runtime packages

The runtime packages contain self-contained .NET executables. They are never
trimmed: trimming turns off the built-in COM support PowerPoint automation needs
(#103). The small launcher
packages select the runtime matching Node's architecture and preserve arguments,
standard streams, signals, and exit codes. Release builds validate all archives;
x64 packages run smoke tests on the Windows x64 runner, and ARM64 packages must
pass a native Windows ARM64 execution gate before the release tag is created.
CI runs the same x64 smoke tests on every pull request. The smoke tests call
`presentation create` (MCP) and `session create` (CLI): without PowerPoint they
must report that PowerPoint is not installed; with PowerPoint they create a
presentation, add a blank slide, and close it. To check a downloaded release
exe on a machine with PowerPoint, run
`node npm-packages/mcp-server-powerpoint/scripts/verify-runtime.mjs <path>\mcp-powerpoint.exe --require-powerpoint`
or the matching `npm-packages/pptcli/scripts/verify-runtime.mjs` with `powerpointcli.exe`.

### Note on the `[Unreleased]` → `[0.0.1]` transition

Before the changesets pipeline existed, `CHANGELOG.md`'s `## [Unreleased]`
section was accumulated by hand. When `v0.0.1` — the first tagged release —
shipped, that section's header was manually renamed to
`## [0.0.1] - 2026-07-09` as a one-time transition step (the changesets
pipeline didn't run for that release). All subsequent releases are fully
changeset-generated via `scripts/Build-Changelog.ps1`; no further manual
`[Unreleased]` handling is needed.
