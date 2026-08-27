"""MkDocs build hook: generate documentation pages from canonical repo sources.

This preserves the project's single-source-of-truth design: several site pages
are generated from the authoritative Markdown files elsewhere in the repo
(CHANGELOG.md, src/*/README.md, skills/README.md) so the website can never
drift from the real docs. It is the MkDocs equivalent of the old Jekyll
``build.sh`` script.

Generated files are written to ``_generated/`` outside the watched docs tree and pulled
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

import json
import logging
import posixpath
import re
from pathlib import Path

log = logging.getLogger("mkdocs.hooks.generate")

# gh-pages/hooks.py -> gh-pages/ -> repo root
REPO_ROOT = Path(__file__).resolve().parent.parent
GEN_DIR = Path(__file__).resolve().parent / "_generated"

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

SKILL_SOURCES = {
    "workflows.md": "Workflows",
    "behavioral-rules.md": "Behavioral Rules",
    "anti-patterns.md": "Anti-Patterns",
    "deck-builder.md": "Deck Builder",
    "slides-and-shapes.md": "Slides and Shapes",
    "tags.md": "String Tags",
    "text-formatting.md": "Text Formatting",
    "tables.md": "Tables",
    "charts.md": "Charts",
    "images.md": "Images",
    "smart-art.md": "SmartArt",
    "speaker-notes.md": "Speaker Notes",
    "layouts.md": "Layouts",
    "master.md": "Slide Masters",
    "animations.md": "Animations",
    "export-and-verify.md": "Export and Verify",
}

SITE_PAGE_MAP.update(
    {
        f"skills/shared/{name}": f"/reference/{Path(name).stem}/"
        for name in SKILL_SOURCES
    }
)

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")
_SNIPPET = re.compile(r'^[ \t]*--8<--[ \t]+"([^"]+)"[ \t]*$', re.MULTILINE)
_FRONTMATTER = re.compile(r"\A---\r?\n.*?\r?\n---\r?\n", re.DOTALL)
_PAGE_MARKDOWN: dict[str, dict[str, str]] = {}
_NAV: list = []
SITE_URL = "https://powerpointmcpserver.dev/"
_FEATURE_HEADLINE = re.compile(
    r"exposes \*\*(?P<tools>\d+) MCP tools with (?P<operations>\d+) operations"
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


def _feature_totals(features: str) -> tuple[int, int]:
    headline = _FEATURE_HEADLINE.search(features)
    if headline is None:
        raise RuntimeError("could not read tool totals from gh-pages/docs/features.md")
    return int(headline.group("tools")), int(headline.group("operations"))


def _write(name: str, source_rel: str, content: str) -> None:
    GEN_DIR.mkdir(parents=True, exist_ok=True)
    content = _rewrite_links(content, source_rel)
    (GEN_DIR / name).write_text(content, encoding="utf-8")
    log.info("generated _generated/%s", name)


def on_pre_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _PAGE_MARKDOWN.clear()
    _NAV.clear()

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

    for name in SKILL_SOURCES:
        _write(
            f"skills-{Path(name).stem}.md",
            f"skills/shared/{name}",
            _strip_header(_read(f"skills/shared/{name}"), demote_h1=True),
        )


DOCS_DIR = Path(__file__).resolve().parent / "docs"
SNIPPET_BASE_PATHS = (DOCS_DIR, Path(__file__).resolve().parent)


def _resolve_snippets(text: str, depth: int = 0) -> str:
    if depth > 5:
        raise RuntimeError("snippet include nesting exceeds five levels")

    def replace(match: re.Match) -> str:
        for base in SNIPPET_BASE_PATHS:
            target = base / match.group(1)
            if target.is_file():
                return _resolve_snippets(target.read_text(encoding="utf-8"), depth + 1)
        raise FileNotFoundError(f"Snippet not found: {match.group(1)}")

    return _SNIPPET.sub(replace, text)


def on_nav(nav, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _NAV.clear()
    _NAV.extend(nav.items)
    return nav


def on_page_markdown(markdown, page, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    body = _resolve_snippets(_FRONTMATTER.sub("", markdown)).strip()
    _PAGE_MARKDOWN[page.file.src_uri] = {
        "title": page.title or page.file.src_uri,
        "url": SITE_URL + page.url,
        "description": (page.meta or {}).get("description", "").strip(),
        "markdown": body,
        "dest": page.file.dest_uri,
    }
    return markdown


def _nav_entries(items, output: list) -> None:
    for item in items:
        if getattr(item, "children", None):
            _nav_entries(item.children, output)
        elif getattr(item, "file", None) is not None:
            output.append(item)


def _write_llm_outputs(config) -> None:
    site_dir = Path(config["site_dir"])
    tool_count, operation_count = _feature_totals(
        _read("gh-pages/docs/features.md")
    )

    for entry in _PAGE_MARKDOWN.values():
        destination = site_dir / entry["dest"]
        if destination.suffix != ".html":
            continue
        markdown_path = destination.with_suffix(".md")
        markdown_path.parent.mkdir(parents=True, exist_ok=True)
        markdown_path.write_text(
            entry["markdown"] + "\n",
            encoding="utf-8",
            newline="\n",
        )

    lines = [
        "# PowerPoint MCP Server",
        "",
        "> PowerPoint MCP Server automates the real Microsoft PowerPoint desktop "
        f"application through its COM API. It exposes {tool_count} tools and "
        f"{operation_count} operations "
        "to AI assistants over the Model Context Protocol and to scripts through "
        "the `pptcli` command line. Windows-only; requires Microsoft PowerPoint.",
        "",
        "Every page below is also available as Markdown by appending `index.md` "
        f"to its URL. The complete corpus is at {SITE_URL}llms-full.txt.",
        "",
    ]

    def link_line(entry: dict) -> str:
        url = entry["url"].rstrip("/")
        if entry["dest"].endswith("index.html"):
            url += "/index.md"
        description = f": {entry['description']}" if entry["description"] else ""
        return f"- [{entry['title']}]({url}){description}"

    seen: set[str] = set()
    for section in _NAV:
        pages: list = []
        _nav_entries([section], pages)
        rendered = []
        for item in pages:
            entry = _PAGE_MARKDOWN.get(item.file.src_uri)
            if entry is None or item.file.src_uri in seen:
                continue
            seen.add(item.file.src_uri)
            rendered.append(link_line(entry))
        if rendered:
            lines.extend([f"## {section.title or 'Documentation'}", "", *rendered, ""])

    (site_dir / "llms.txt").write_text(
        "\n".join(lines),
        encoding="utf-8",
        newline="\n",
    )

    ordered: list = []
    _nav_entries(_NAV, ordered)
    full = ["# PowerPoint MCP Server - complete documentation", ""]
    emitted: set[str] = set()
    for item in ordered:
        entry = _PAGE_MARKDOWN.get(item.file.src_uri)
        if entry is None or item.file.src_uri in emitted:
            continue
        emitted.add(item.file.src_uri)
        full.extend(
            [
                f"# {entry['title']}",
                "",
                f"Source: {entry['url']}",
                "",
                entry["markdown"],
                "",
                "---",
                "",
            ]
        )

    (site_dir / "llms-full.txt").write_text(
        "\n".join(full),
        encoding="utf-8",
        newline="\n",
    )
    log.info("wrote llms.txt, llms-full.txt and %d Markdown mirrors", len(emitted))


def _write_tools_json(config) -> None:
    features = _read("gh-pages/docs/features.md")
    expected_tools, expected_operations = _feature_totals(features)

    matrix = {}
    for match in re.finditer(
        r"^\| `(?P<name>[^`]+)` \| (?P<count>\d+) \| (?P<description>[^|]+) \|",
        features,
        re.MULTILINE,
    ):
        matrix[match.group("name")] = {
            "operationCount": int(match.group("count")),
            "description": match.group("description").strip(),
        }

    heading = re.compile(
        r"^### `(?P<name>[^`]+)` tool \((?P<count>\d+) operations\)$",
        re.MULTILINE,
    )
    matches = list(heading.finditer(features))
    tools = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(features)
        section = features[match.end() : end]
        action_block = re.search(
            r"\*\*Exact action order:\*\*(?P<actions>.*?)(?:\n\s*\n|\Z)",
            section,
            re.DOTALL,
        )
        if action_block is None:
            raise RuntimeError(f"missing exact action order for {match.group('name')}")
        actions = re.findall(r"`([^`]+)`", action_block.group("actions"))
        expected = int(match.group("count"))
        if len(actions) != expected:
            raise RuntimeError(
                f"{match.group('name')} documents {len(actions)} actions, expected {expected}"
            )
        matrix_entry = matrix.get(match.group("name"))
        if matrix_entry is None or matrix_entry["operationCount"] != expected:
            raise RuntimeError(f"tool matrix is missing or stale for {match.group('name')}")
        tools.append(
            {
                "name": match.group("name"),
                "description": matrix_entry["description"],
                "operationCount": expected,
                "operations": [{"name": action} for action in actions],
            }
        )

    actual_operations = sum(tool["operationCount"] for tool in tools)
    if len(tools) != expected_tools or actual_operations != expected_operations:
        raise RuntimeError(
            "tools.json totals do not match the feature headline: "
            f"{len(tools)} tools/{actual_operations} operations vs "
            f"{expected_tools}/{expected_operations}"
        )

    payload = {
        "name": "PowerPoint MCP Server",
        "url": SITE_URL,
        "repository": "https://github.com/sbroenne/mcp-server-powerpoint",
        "description": (
            "Automates the real Microsoft PowerPoint desktop application through "
            "its COM API for AI assistants and command-line scripts."
        ),
        "requirements": {
            "operatingSystem": "Windows",
            "application": "Microsoft PowerPoint desktop",
        },
        "entryPoints": ["mcp-server", "cli"],
        "toolCount": len(tools),
        "operationCount": actual_operations,
        "categories": tools,
    }
    (Path(config["site_dir"]) / "tools.json").write_text(
        json.dumps(payload, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    log.info("wrote tools.json (%d tools, %d operations)", len(tools), actual_operations)


def on_post_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _write_llm_outputs(config)
    _write_tools_json(config)


def on_post_page(output, page, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Give Material's search dialog an accessible name."""
    output = output.replace(
        '<div class="md-search" data-md-component="search" role="dialog">',
        '<div class="md-search" data-md-component="search" role="dialog" '
        'aria-label="Search documentation">',
    )
    output = output.replace(
        "<div class=md-search data-md-component=search role=dialog>",
        '<div class=md-search data-md-component=search role=dialog '
        'aria-label="Search documentation">',
    )
    return output
