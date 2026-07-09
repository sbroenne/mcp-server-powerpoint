"""MkDocs build hook: generate documentation pages from canonical repo sources.

This preserves the project's single-source-of-truth design: several site pages
are generated from the authoritative Markdown files elsewhere in the repo
(CHANGELOG.md, src/*/README.md, skills/README.md) so the website can never
drift from the real docs. It is the MkDocs equivalent of the old Jekyll
``build.sh`` script.

Generated files are written to ``docs/_generated/`` (git-ignored) and pulled
into the thin wrapper pages under ``docs/`` via the ``pymdownx.snippets``
``--8<--`` include syntax. Regeneration happens automatically on every
``mkdocs build`` / ``mkdocs serve`` via the ``on_pre_build`` event.

Unlike the sibling mcp-server-excel repo, this repo does not yet have a root
FEATURES.md or a docs/INSTALLATION.md / CONTRIBUTING.md / SECURITY.md /
PRIVACY.md set — those site pages (features, installation, contributing,
security, privacy) remain hand-authored directly under gh-pages/docs/ until
such root docs exist.
"""

from __future__ import annotations

import logging
import posixpath
import re
from pathlib import Path

log = logging.getLogger("mkdocs.hooks.generate")

# gh-pages/hooks.py -> gh-pages/ -> repo root
REPO_ROOT = Path(__file__).resolve().parent.parent
GEN_DIR = Path(__file__).resolve().parent / "docs" / "_generated"

GITHUB_BLOB = "https://github.com/sbroenne/mcp-server-powerpoint/blob/main/"
GITHUB_TREE = "https://github.com/sbroenne/mcp-server-powerpoint/tree/main/"

# Repo-relative paths that have a dedicated site page: rewrite links to them so
# they resolve on the website instead of 404-ing.
SITE_PAGE_MAP = {
    "CHANGELOG.md": "/changelog/",
    "src/PowerPointMcp.McpServer/README.md": "/mcp-server/",
    "src/PowerPointMcp.CLI/README.md": "/cli/",
    "skills/README.md": "/skills/",
}

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")


def _rewrite_links(text: str, source_rel: str) -> str:
    """Resolve repo-relative links in pulled-in content so they work on the site.

    Links that point at a page we publish are rewritten to that page's URL;
    everything else that resolves inside the repo is rewritten to an absolute
    GitHub URL. External links, anchors and site-absolute links are left alone.
    """
    source_dir = posixpath.dirname(source_rel)

    def repl(match: re.Match) -> str:
        label, url = match.group(1), match.group(2)
        if url.startswith(("http://", "https://", "#", "/", "mailto:", "<")):
            return match.group(0)

        anchor = ""
        target = url
        if "#" in target:
            target, anchor = target.split("#", 1)
            anchor = "#" + anchor
        if target == "":
            return match.group(0)  # pure in-page anchor

        resolved = posixpath.normpath(posixpath.join(source_dir, target))
        if resolved.startswith(".."):
            return match.group(0)  # points outside the repo; leave as-is

        if resolved in SITE_PAGE_MAP:
            return f"[{label}]({SITE_PAGE_MAP[resolved]}{anchor})"

        base = GITHUB_TREE if url.endswith("/") else GITHUB_BLOB
        return f"[{label}]({base}{resolved}{anchor})"

    return _MD_LINK.sub(repl, text)


def _strip_header(text: str, *, demote_h1: bool = False) -> str:
    """Drop the leading H1 title block (and any badges/description lines
    beneath it) from a source file, up to but not including the next heading.

    Mirrors the awk transforms in the previous Jekyll ``build.sh``: the first
    ``# Title`` line and everything below it (badges, one-line descriptions,
    blank lines) are dropped until the next Markdown heading is reached, since
    that heading marks the start of real content. When ``demote_h1`` is set,
    any later top-level ``# `` heading found in the remaining content is
    demoted to ``## `` so it nests correctly under the site's own page title.
    """
    lines = text.splitlines()
    start = 0
    seen_title = False
    for i, line in enumerate(lines):
        if not seen_title:
            if line.startswith("# "):
                seen_title = True
            continue
        if line.startswith("#"):
            start = i
            break
    else:
        start = len(lines) if seen_title else 0

    out: list[str] = []
    for line in lines[start:]:
        if demote_h1 and line.startswith("# "):
            line = "#" + line  # "# " -> "## "
        out.append(line)

    return "\n".join(out).strip() + "\n"


def _read(rel: str) -> str:
    path = REPO_ROOT / rel
    if not path.is_file():
        raise FileNotFoundError(f"Source doc not found: {path}")
    return path.read_text(encoding="utf-8")


def _write(name: str, source_rel: str, content: str) -> None:
    GEN_DIR.mkdir(parents=True, exist_ok=True)
    content = _rewrite_links(content, source_rel)
    (GEN_DIR / name).write_text(content, encoding="utf-8")
    log.info("generated _generated/%s", name)


def on_pre_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    # CHANGELOG.md -> changelog (drop title + description paragraph, demote H1)
    _write(
        "changelog.md",
        "CHANGELOG.md",
        _strip_header(_read("CHANGELOG.md"), demote_h1=True),
    )

    # src/PowerPointMcp.McpServer/README.md -> mcp-server (drop title + badges)
    _write(
        "mcp-server.md",
        "src/PowerPointMcp.McpServer/README.md",
        _strip_header(_read("src/PowerPointMcp.McpServer/README.md"), demote_h1=True),
    )

    # src/PowerPointMcp.CLI/README.md -> cli (drop title + badges, demote H1)
    _write(
        "cli.md",
        "src/PowerPointMcp.CLI/README.md",
        _strip_header(_read("src/PowerPointMcp.CLI/README.md"), demote_h1=True),
    )

    # skills/README.md -> skills (drop title, demote H1)
    _write(
        "skills.md",
        "skills/README.md",
        _strip_header(_read("skills/README.md"), demote_h1=True),
    )
