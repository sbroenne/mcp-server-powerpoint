# Release Strategy

PowerPointMcp releases all components together (MCP Server, CLI, VS Code Extension,
MCPB, Agent Skills, MCP Registry entry) under a single version number, via the
`.github/workflows/release.yml` `workflow_dispatch` workflow.

## Cutting a release

1. Trigger `Release All Components` from the Actions tab with a `version_bump`
   (major/minor/patch) or a `custom_version`.
2. The workflow calculates the next version from the latest git tag, builds and
   publishes every component (NuGet, standalone exe zips, VS Code Marketplace,
   MCPB, Agent Skills zip, MCP Registry), creates the git tag, then creates the
   GitHub Release.

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
  updated `CHANGELOG.md`, `package.json`, and consumed `.changeset/*.md`
  deletions back to `main`, since branch protection prevents a direct push.
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

`package.json` and `.changeset/` exist **solely** to host the `@changesets/cli`
tool for `CHANGELOG.md` generation. Node is not used to build, version, or publish
any PowerPointMcp component — the MCP Server, CLI, VS Code Extension, and MCPB
all remain built by their existing pipelines (`dotnet`, `vsce`, etc.).

### Note on the pre-1.0 `[Unreleased]` section

Since this repo hasn't cut its first tagged release yet, `CHANGELOG.md`'s
existing `## [Unreleased]` section predates the changesets pipeline and was
accumulated by hand. When cutting the first release, manually rename that
section's header to `## [<version>] - <date>` as part of the release PR
*before* running `scripts/Build-Changelog.ps1` for any subsequent release —
after that one-time transition, all future entries are changeset-generated.
