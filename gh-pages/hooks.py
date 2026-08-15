"""MkDocs build hook: generate documentation pages from canonical repo sources.

This preserves the project's single-source-of-truth design: many site pages are
generated from the authoritative Markdown files elsewhere in the repo
(CHANGELOG.md, src/*/README.md, skills/README.md, docs/guides/*.md and the
skills/shared/*.md reference corpus) so the website can never drift from the
real docs. It is the MkDocs equivalent of the old Jekyll ``build.sh`` script.

Generated files are written to ``docs/_generated/`` (git-ignored) and pulled
into the thin wrapper pages under ``docs/`` via the ``pymdownx.snippets``
``--8<--`` include syntax. The wrappers own presentation and SEO metadata only;
they never duplicate prose. Regeneration happens automatically on every
``mkdocs build`` / ``mkdocs serve`` via the ``on_pre_build`` event.

Unlike the sibling mcp-server-excel repo, this repo does not yet have a root
FEATURES.md or a docs/INSTALLATION.md / CONTRIBUTING.md / SECURITY.md /
PRIVACY.md set — those site pages (features, installation, contributing,
security, privacy) remain hand-authored directly under gh-pages/docs/ until
such root docs exist.
"""

from __future__ import annotations

import logging
import json
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
    "docs/guides/README.md": "/guides/",
    "docs/guides/BUILD-A-DECK-WITH-AI.md": "/guides/build-a-deck-with-ai/",
    "docs/guides/EDIT-AN-EXISTING-DECK.md": "/guides/edit-an-existing-deck/",
    "docs/guides/EXPORT-SLIDES-TO-IMAGES.md": "/guides/export-slides-to-images/",
    "docs/guides/AUTOMATE-CHARTS-AND-TABLES.md": "/guides/automate-charts-and-tables/",
    "docs/guides/APPLY-A-CORPORATE-TEMPLATE.md": "/guides/apply-a-corporate-template/",
}

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")

# Canonical task guides -> intent-focused website pages. The repo copy stays
# useful on GitHub; the wrapper page under docs/guides/ owns presentation and
# SEO metadata only and never duplicates the prose.
GUIDE_SOURCES = {
    "guides-index.md": "docs/guides/README.md",
    "guides-build-a-deck-with-ai.md": "docs/guides/BUILD-A-DECK-WITH-AI.md",
    "guides-edit-an-existing-deck.md": "docs/guides/EDIT-AN-EXISTING-DECK.md",
    "guides-export-slides-to-images.md": "docs/guides/EXPORT-SLIDES-TO-IMAGES.md",
    "guides-automate-charts-and-tables.md": "docs/guides/AUTOMATE-CHARTS-AND-TABLES.md",
    "guides-apply-a-corporate-template.md": "docs/guides/APPLY-A-CORPORATE-TEMPLATE.md",
}

# skills/shared/*.md: the expert reference corpus shipped inside the skill
# packages. Published verbatim so the website and the agent guidance can never
# disagree. Value = (generated file name, url slug).
SKILL_SOURCES = {
    "workflows.md": ("skills-workflows.md", "workflows"),
    "behavioral-rules.md": ("skills-behavioral-rules.md", "behavioral-rules"),
    "anti-patterns.md": ("skills-anti-patterns.md", "anti-patterns"),
    "deck-builder.md": ("skills-deck-builder.md", "deck-builder"),
    "slides-and-shapes.md": ("skills-slides-and-shapes.md", "slides-and-shapes"),
    "text-formatting.md": ("skills-text-formatting.md", "text-formatting"),
    "tables.md": ("skills-tables.md", "tables"),
    "charts.md": ("skills-charts.md", "charts"),
    "images.md": ("skills-images.md", "images"),
    "smart-art.md": ("skills-smart-art.md", "smartart"),
    "animations.md": ("skills-animations.md", "animations"),
    "speaker-notes.md": ("skills-speaker-notes.md", "speaker-notes"),
    "layouts.md": ("skills-layouts.md", "layouts"),
    "master.md": ("skills-master.md", "slide-master"),
    "export-and-verify.md": ("skills-export-and-verify.md", "export-and-verify"),
}

SITE_PAGE_MAP.update(
    {f"skills/shared/{name}": f"/reference/{slug}/" for name, (_out, slug) in SKILL_SOURCES.items()}
)


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


def _strip_header(text: str, *, demote_h1: bool = False, end_on_blank: bool = False) -> str:
    """Drop the leading H1 title block (and any badges/description lines
    beneath it) from a source file, up to but not including the next heading.

    Mirrors the awk transforms in the previous Jekyll ``build.sh``: the first
    ``# Title`` line and everything below it (badges, one-line descriptions,
    blank lines) are dropped until the next Markdown heading is reached, since
    that heading marks the start of real content. When ``demote_h1`` is set,
    any later top-level ``# `` heading found in the remaining content is
    demoted to ``## `` so it nests correctly under the site's own page title.

    ``end_on_blank`` stops the drop at the first blank line after the title
    instead of running on to the next heading. Use it for sources whose H1 is
    followed by an intro paragraph worth keeping (the wrapper page supplies the
    title, but the prose is real content).
    """
    lines = text.splitlines()
    start = 0
    seen_title = False
    for i, line in enumerate(lines):
        if not seen_title:
            if line.startswith("# "):
                seen_title = True
            continue
        if end_on_blank and not line.strip():
            start = i + 1
            break
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


# The real star-history chart is produced by scripts/Update-StarHistory.ps1 in the
# deploy workflow and downloaded into docs/assets/images/ before mkdocs runs. It
# is never committed, so on a developer machine the link in index.md dangles and
# `mkdocs build --strict` fails on a warning that says nothing about the change
# being made. Drop in a neutral placeholder when it is missing; CI always has the
# real file by this point, so this only ever fires locally.
STAR_HISTORY_SVG = Path(__file__).resolve().parent / "docs" / "assets" / "images" / "star-history.svg"

_STAR_HISTORY_PLACEHOLDER = """<svg xmlns="http://www.w3.org/2000/svg" width="720" height="360" \
viewBox="0 0 720 360" role="img" aria-label="Star history chart placeholder">
  <rect width="720" height="360" fill="#f5f7fa"/>
  <text x="360" y="180" text-anchor="middle" dominant-baseline="middle"
        font-family="sans-serif" font-size="18" fill="#7a869a">
    Star history chart is generated during deployment
  </text>
</svg>
"""


def _ensure_star_history_placeholder() -> None:
    if STAR_HISTORY_SVG.exists():
        return
    STAR_HISTORY_SVG.parent.mkdir(parents=True, exist_ok=True)
    STAR_HISTORY_SVG.write_text(_STAR_HISTORY_PLACEHOLDER, encoding="utf-8")
    log.info("wrote placeholder star-history.svg (real chart is generated in CI)")


def on_pre_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _ensure_star_history_placeholder()

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

    # Canonical task guides -> intent-focused website pages. The H1 lives in the
    # wrapper, so drop it here but keep the intro paragraph beneath it.
    for output_name, source_rel in GUIDE_SOURCES.items():
        _write(
            output_name,
            source_rel,
            _strip_header(_read(source_rel), demote_h1=True, end_on_blank=True),
        )

    # skills/shared/*.md -> reference pages (drop the H1, wrapper owns the title)
    for name, (output_name, _slug) in SKILL_SOURCES.items():
        _write(
            output_name,
            f"skills/shared/{name}",
            _strip_header(
                _read(f"skills/shared/{name}"), demote_h1=True, end_on_blank=True
            ),
        )


# ---------------------------------------------------------------------------
# LLM-facing discoverability layer
#
# Emits, at the end of every build:
#   * /llms.txt       - llmstxt.org-style index, ordered by the resolved nav
#   * /llms-full.txt  - the whole corpus as one Markdown file
#   * <page>/index.md - a Markdown mirror beside every generated index.html
#
# All three derive from the same captured page Markdown, so they cannot drift
# from the rendered site. Ported from the sibling mcp-server-excel repo.
# ---------------------------------------------------------------------------

SITE_URL = "https://powerpointmcpserver.dev/"

_FRONTMATTER = re.compile(r"\A---\n.*?\n---\n", re.DOTALL)
_SNIPPET = re.compile(r"^--8<--\s+\"([^\"]+)\"\s*$", re.MULTILINE)

# src_uri -> captured page data, filled in by on_page_markdown.
_PAGE_MARKDOWN: dict[str, dict] = {}

# The resolved Navigation object, captured in on_nav. config["nav"] holds the raw
# YAML nav, which has no page objects to correlate with captured Markdown.
_NAV: list = []


def _resolve_snippets(text: str, depth: int = 0) -> str:
    """Inline `--8<-- "..."` includes so the LLM outputs carry real content.

    pymdownx.snippets runs after this hook, so the captured Markdown would
    otherwise contain include directives instead of the generated docs.
    """
    if depth > 5:
        return text

    def repl(match: re.Match) -> str:
        target = Path(__file__).resolve().parent / "docs" / match.group(1)
        if not target.is_file():
            log.warning("snippet not found while building llms output: %s", target)
            return ""
        return _resolve_snippets(target.read_text(encoding="utf-8"), depth + 1)

    return _SNIPPET.sub(repl, text)


_FAQ_ADMONITION = re.compile(r'^\?{3}\+?\s+question\s+"([^"]+)"\s*$')
_FAQ_HEADING = re.compile(r"^###\s+(.+?)\s*$")


def _faq_jsonld(markdown: str) -> str:
    """Build FAQPage JSON-LD from a page's own question blocks.

    Two source forms are recognised:

    * ``### Some question?`` headings — preferred, because each question keeps a
      stable anchor that can be deep-linked from other pages and from search
      results. Only headings that read as questions are collected.
    * ``??? question "..."`` admonitions, matching the sibling mcp-server-excel
      repo, so a page written either way works.

    Either way the structured data is derived from the rendered body rather than
    maintained separately, so the two cannot diverge.
    """
    items: list[tuple[str, list[str]]] = []
    current: list[str] | None = None
    indented = False

    for line in markdown.splitlines():
        admonition = _FAQ_ADMONITION.match(line)
        if admonition:
            current = []
            indented = True
            items.append((admonition.group(1), current))
            continue

        heading = _FAQ_HEADING.match(line)
        if heading:
            text = heading.group(1).strip()
            if text.endswith("?"):
                current = []
                indented = False
                items.append((text, current))
            else:
                current = None
            continue

        if current is None:
            continue

        # A new heading of any level ends a heading-sourced answer.
        if not indented and line.startswith("#"):
            current = None
            continue

        if not line.strip():
            current.append("")
        elif indented and not line.startswith((" ", "\t")):
            current = None
        else:
            current.append(line.strip())

    entities = []
    for question, answer_lines in items:
        # Drop fenced code blocks and table rows: useful on the page, but noise
        # in a structured answer. (Improvement over the Excel implementation,
        # which joins every captured line verbatim.)
        prose: list[str] = []
        in_fence = False
        for raw in answer_lines:
            if raw.startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence or raw.startswith("|"):
                continue
            prose.append(raw)

        answer = " ".join(x for x in prose if x).strip()
        if not answer:
            continue
        # Strip inline Markdown so the structured answer is plain prose.
        answer = _MD_LINK.sub(r"\1", answer)
        answer = re.sub(r"[*_`]+", "", answer)
        answer = re.sub(r"(?:^|\s)-\s+", " ", answer)
        answer = re.sub(r"\s{2,}", " ", answer).strip()
        entities.append(
            {
                "@type": "Question",
                "name": question,
                "acceptedAnswer": {"@type": "Answer", "text": answer},
            }
        )

    if not entities:
        return ""

    return json.dumps(
        {"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": entities},
        ensure_ascii=False,
    )


def on_nav(nav, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _NAV.clear()
    _NAV.extend(nav.items)
    return nav


def on_page_markdown(markdown, page, config, **kwargs):  # noqa: D401 - MkDocs hook
    """Capture each page's full Markdown for the LLM-facing outputs."""
    body = _resolve_snippets(_FRONTMATTER.sub("", markdown)).strip()
    _PAGE_MARKDOWN[page.file.src_uri] = {
        "title": page.title or page.file.src_uri,
        "url": SITE_URL + page.url,
        "description": (page.meta or {}).get("description", "").strip(),
        "markdown": body,
        "dest": page.file.dest_uri,
    }

    faq = _faq_jsonld(body)
    if faq:
        page.meta["faq_jsonld"] = faq
    return markdown


def _nav_entries(items, out: list) -> None:
    for item in items:
        if getattr(item, "children", None):
            _nav_entries(item.children, out)
        elif getattr(item, "file", None) is not None:
            out.append(item)


def on_post_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Emit /llms.txt, /llms-full.txt and one Markdown mirror per page."""
    site_dir = Path(config["site_dir"])

    # Markdown mirrors: /faq/index.md sits next to /faq/index.html.
    mirrored = 0
    for entry in _PAGE_MARKDOWN.values():
        dest = site_dir / entry["dest"]
        if dest.suffix != ".html":
            continue
        md_path = dest.with_suffix(".md")
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text(entry["markdown"] + "\n", encoding="utf-8", newline="\n")
        mirrored += 1

    # Section-aware index, ordered exactly like the site navigation.
    lines = [
        "# PowerPoint MCP Server",
        "",
        "> PowerPoint MCP Server (PowerPointMcp) automates the real Microsoft "
        "PowerPoint application through its COM API, exposing 13 tools and 141 "
        "operations to AI assistants over the Model Context Protocol and to "
        "scripts through the `pptcli` command line. Unlike file-parser libraries "
        "such as python-pptx it drives a live PowerPoint instance, so it can "
        "render SmartArt, apply real themes and layouts, run animations, and "
        "export slides to images for visual verification. Windows-only; requires "
        "Microsoft PowerPoint 2016 or later.",
        "",
        "Every page below is also available as Markdown by appending `index.md` "
        f"to its URL. The complete corpus is at {SITE_URL}llms-full.txt.",
        "",
    ]

    def link_line(entry: dict) -> str:
        url = entry["url"].rstrip("/")
        url = f"{url}/index.md" if entry["dest"].endswith("index.html") else url
        desc = f": {entry['description']}" if entry["description"] else ""
        return f"- [{entry['title']}]({url}){desc}"

    seen: set[str] = set()
    for section in _NAV:
        pages: list = []
        _nav_entries([section], pages)
        title = section.title if getattr(section, "title", None) else "Documentation"
        rendered = []
        for item in pages:
            entry = _PAGE_MARKDOWN.get(item.file.src_uri)
            if entry is None or item.file.src_uri in seen:
                continue
            seen.add(item.file.src_uri)
            rendered.append(link_line(entry))
        if rendered:
            lines.append(f"## {title}")
            lines.append("")
            lines.extend(rendered)
            lines.append("")

    (site_dir / "llms.txt").write_text("\n".join(lines), encoding="utf-8", newline="\n")

    # Full corpus, same order as llms.txt.
    full = ["# PowerPoint MCP Server - complete documentation", ""]
    ordered: list = []
    _nav_entries(_NAV, ordered)
    emitted: set[str] = set()
    for item in ordered:
        entry = _PAGE_MARKDOWN.get(item.file.src_uri)
        if entry is None or item.file.src_uri in emitted:
            continue
        emitted.add(item.file.src_uri)
        full.append(f"# {entry['title']}")
        full.append("")
        full.append(f"Source: {entry['url']}")
        full.append("")
        full.append(entry["markdown"])
        full.append("")
        full.append("---")
    full.append("")
    (site_dir / "llms-full.txt").write_text(
        "\n".join(full), encoding="utf-8", newline="\n"
    )

    log.info("wrote llms.txt, llms-full.txt and %d Markdown mirrors", mirrored)
