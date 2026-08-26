# Docs Site (MkDocs)

Source for [powerpointmcpserver.dev](https://powerpointmcpserver.dev/), built with MkDocs
Material. Several pages under `docs/` are thin wrappers that pull in canonical content from
elsewhere in the repo (`CHANGELOG.md`, `src/PowerPointMcp.McpServer/README.md`,
`src/PowerPointMcp.CLI/README.md`, `skills/README.md`) via `hooks.py` and the
`pymdownx.snippets` `--8<--` include syntax, so there is a single source of truth for that
content. Other pages (home, features, installation, contributing, security, privacy) are
hand-authored directly under `docs/`.

## Setup (one-time)

```powershell
cd gh-pages
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## ⚠️ Always use the venv Python

A global `mkdocs` on `PATH` may resolve to a different Python install that does not have
`mkdocs-material` (or a stale version of it). Always invoke mkdocs through the project's venv:

```powershell
cd gh-pages
.\.venv\Scripts\python.exe -m mkdocs serve   # live preview with auto-reload
.\.venv\Scripts\python.exe -m mkdocs build --strict --clean   # verify before commit
```

(Alternatively, activate the venv first with `.\.venv\Scripts\Activate.ps1`, then plain
`mkdocs serve`/`mkdocs build` will use the correct interpreter.)

## How generated pages work

`hooks.py` regenerates `_generated/*.md` from canonical repo sources on every build
(`on_pre_build`). Keeping generated files outside `docs/` prevents MkDocs serve from rebuilding
in a loop. The directory is git-ignored — never edit it directly. Edit the source file instead
(e.g. `CHANGELOG.md` at the repo root) and rebuild.

## Checks

The `Docs Site` CI job builds with strict warnings, audits internal links, and verifies that the
Pages workflow watches every canonical source mirrored by `hooks.py`. Run the same checks locally
after a build:

```powershell
cd gh-pages
.\.venv\Scripts\python.exe -m mkdocs build --strict --clean
.\.venv\Scripts\python.exe audit_site.py
.\.venv\Scripts\python.exe check_deploy_paths.py
```

The weekly `link-check.yml` workflow checks external links and files an issue when links break.
It intentionally does not run on pull requests because external sites can rate-limit requests or
be temporarily unavailable.
