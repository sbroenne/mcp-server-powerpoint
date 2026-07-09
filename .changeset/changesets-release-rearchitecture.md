---
"powerpointmcp": patch
---

Replaced manual `CHANGELOG.md` editing and the fragile awk/sed-based release-notes
extraction with [changesets](https://github.com/changesets/changesets): contributors
now add a small `.changeset/*.md` fragment describing their change, CI enforces one
is present (or the `skip-changelog` label), and `scripts/Build-Changelog.ps1`
compiles pending fragments into `CHANGELOG.md` and the GitHub Release body at
release time. See `docs/RELEASE-STRATEGY.md` for the full process.
