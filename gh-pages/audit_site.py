#!/usr/bin/env python3
"""Audit the built MkDocs site for SEO and LLM-discoverability regressions.

Run after ``mkdocs build`` from the ``gh-pages`` directory::

    python -m mkdocs build --strict --clean
    python audit_site.py

Exits non-zero and prints every failure if the built site regresses. This runs
in the Pages deploy workflow rather than the pre-commit hook, so the docs-only
pre-commit fast path stays fast.

Ported from the sibling mcp-server-excel repo. Note that the ``minify`` plugin
strips attribute quotes, so every attribute regex here must tolerate both
``type="x"`` and ``type=x``.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import urlsplit

SITE_DIR = Path(__file__).resolve().parent / "_site"
SITE_URL = "https://powerpointmcpserver.dev/"

# mkdocs.yml's site_description. Material falls back to it whenever a page has no
# usable per-page description, so a page rendering it verbatim means that page's
# YAML front matter never applied - and MkDocs reports nothing when that happens.
# Two ways to trigger it, both found in this repo: a double quote in the value
# (which terminates the HTML content="..." attribute early) and an unquoted YAML
# scalar containing ": " (which makes the whole front-matter block unparseable,
# silently dropping title, description and keywords together).
#
# Read from mkdocs.yml rather than duplicated here: a hardcoded copy would stop
# matching the moment site_description was reworded, and this check would then
# silently pass forever without guarding anything.
def _site_description() -> str:
    """Read site_description from mkdocs.yml, handling plain and block scalars.

    mkdocs.yml carries custom tags that ``yaml.safe_load`` rejects, so this
    parses the one key it needs rather than pulling in a full YAML load.
    """
    path = Path(__file__).resolve().parent / "mkdocs.yml"
    lines = path.read_text(encoding="utf-8").splitlines()

    for index, line in enumerate(lines):
        if not line.startswith("site_description:"):
            continue
        value = line.split(":", 1)[1].strip()

        # Folded (>) or literal (|) block scalar: the value is the indented
        # block that follows, not the indicator itself.
        if value[:1] in {">", "|"}:
            collected: list[str] = []
            for candidate in lines[index + 1 :]:
                if candidate.strip() and not candidate.startswith((" ", "\t")):
                    break
                collected.append(candidate.strip())
            value = " ".join(collected)
        elif value[:1] in {'"', "'"} and value[-1:] == value[:1]:
            value = value[1:-1]

        value = " ".join(value.split())
        if value:
            return value
        break

    print("audit: could not read site_description from mkdocs.yml", file=sys.stderr)
    sys.exit(1)


SITE_DESCRIPTION = _site_description()

# The homepage legitimately uses the site description; every other page must not.
HOMEPAGE = "index.html"

# 404.html is not an indexable page: MkDocs renders it without a canonical URL
# or page metadata, so every metadata check would fire on it.

# Google truncates around these lengths; well outside them is a real problem.
TITLE_MAX = 70
DESCRIPTION_MIN = 50
DESCRIPTION_MAX = 200

# AI crawlers that robots.txt must name explicitly. Being quoted accurately by
# an assistant is a goal of this documentation, so a silent removal is a bug.
REQUIRED_AI_AGENTS = [
    "GPTBot",
    "OAI-SearchBot",
    "ChatGPT-User",
    "ClaudeBot",
    "Claude-User",
    "PerplexityBot",
    "Google-Extended",
]

failures: list[str] = []
checked = 0


def fail(message: str) -> None:
    failures.append(message)


def page_name(path: Path) -> str:
    return path.relative_to(SITE_DIR).as_posix()


def attr(html: str, pattern: str) -> str | None:
    match = re.search(pattern, html, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else None


def audit_html(path: Path) -> None:
    global checked
    checked += 1
    html = path.read_text(encoding="utf-8", errors="replace")
    name = page_name(path)

    canonical = re.search(r'<link[^>]+rel=["\']?canonical["\']?[^>]*>', html, re.I)
    if not canonical:
        fail(f"{name}: no canonical link")
    elif SITE_URL not in canonical.group(0):
        fail(f"{name}: canonical does not point at {SITE_URL}")

    titles = re.findall(r"<title>(.*?)</title>", html, re.DOTALL)
    if not titles:
        fail(f"{name}: no <title>")
    elif len(titles[0].strip()) > TITLE_MAX:
        fail(f"{name}: <title> is {len(titles[0].strip())} chars (max {TITLE_MAX})")

    description = attr(
        html, r'<meta[^>]+name=["\']?description["\']?[^>]+content=["\'](.*?)["\']'
    )
    if description is None:
        fail(f"{name}: no meta description")
    else:
        # Normalise once: the length check and the site_description comparison
        # below must judge the same string, or a page can be reported with a
        # length that does not match the value actually being compared.
        text = " ".join(description.split())
        length = len(text)
        if not DESCRIPTION_MIN <= length <= DESCRIPTION_MAX:
            fail(
                f"{name}: meta description is {length} chars "
                f"(want {DESCRIPTION_MIN}-{DESCRIPTION_MAX})"
            )
        if name != HOMEPAGE and text == SITE_DESCRIPTION:
            fail(
                f"{name}: meta description fell back to site_description - this "
                f"page's YAML front matter did not apply (check for a double quote "
                f'or an unquoted ": " in the description value)'
            )

    # Exactly one H1 per page: multiple H1s dilute the topical signal.
    h1s = re.findall(r"<h1[ >]", html, re.I)
    if len(h1s) != 1:
        fail(f"{name}: expected exactly 1 <h1>, found {len(h1s)}")

    # Social preview completeness.
    for prop in ("og:title", "og:description", "og:url", "og:image"):
        if not re.search(rf'property=["\']?{re.escape(prop)}["\']?', html, re.I):
            fail(f"{name}: missing {prop}")
    if not re.search(r'name=["\']?twitter:card["\']?', html, re.I):
        fail(f"{name}: missing twitter:card")

    # Markdown alternate: the machine-readable mirror of this page.
    if not re.search(r'rel=["\']?alternate["\']?[^>]+text/markdown', html, re.I):
        fail(f"{name}: no rel=alternate Markdown link")

    # Explicit intrinsic size prevents layout shift. Exempt SVGs (they carry
    # their own intrinsic dimensions), the theme logo (Material emits it from
    # its own partials and sizes it via CSS), and remote status badges, whose
    # dimensions are not knowable at build time.
    badge_hosts = ("img.shields.io", "vsmarketplacebadges.dev", "badgen.net")
    for img in re.findall(r"<img[^>]*>", html, re.I):
        src = attr(img, r'src=["\']?([^"\'\s>]+)') or ""
        parts = urlsplit(src)
        if parts.netloc in badge_hosts:
            continue
        if parts.path.endswith(".svg") or parts.path.endswith("/icon.png"):
            continue
        if not re.search(r"\bwidth=", img, re.I) or not re.search(
            r"\bheight=", img, re.I
        ):
            fail(f"{name}: <img> without width/height: {src or img[:60]}")


def audit_internal_links(html_files: list[Path]) -> None:
    """Catch internal links that point at a page the build never produced."""
    for path in html_files:
        html = path.read_text(encoding="utf-8", errors="replace")
        name = page_name(path)
        for href in re.findall(r'<a[^>]+href=["\']?([^"\'\s>]+)', html, re.I):
            if href.startswith(("http://", "https://", "mailto:", "#", "javascript:")):
                continue
            target = urlsplit(href).path
            if not target:
                continue
            resolved = (path.parent / target).resolve()
            if resolved.is_dir():
                resolved = resolved / "index.html"
            elif resolved.suffix == "":
                resolved = resolved.with_suffix(".html")
            try:
                resolved.relative_to(SITE_DIR.resolve())
            except ValueError:
                continue
            if not resolved.exists():
                fail(f"{name}: broken internal link -> {href}")


def audit_sitemap() -> None:
    sitemap = SITE_DIR / "sitemap.xml"
    if not sitemap.exists():
        fail("sitemap.xml is missing")
        return
    text = sitemap.read_text(encoding="utf-8")
    locs = re.findall(r"<loc>(.*?)</loc>", text)
    if not locs:
        fail("sitemap.xml contains no <loc> entries")
    if SITE_URL not in locs:
        fail("sitemap.xml is missing the homepage entry")
    for loc in locs:
        if not loc.startswith(SITE_URL):
            fail(f"sitemap.xml has an off-site <loc>: {loc}")
    if not (SITE_DIR / "sitemap.xml.gz").exists():
        fail("sitemap.xml.gz is missing")


def audit_llms(html_files: list[Path]) -> None:
    index = SITE_DIR / "llms.txt"
    if not index.exists():
        fail("llms.txt is missing")
    else:
        text = index.read_text(encoding="utf-8")
        lines = text.splitlines()
        if not lines or not lines[0].startswith("# "):
            fail("llms.txt must start with an H1")
        if "> " not in text:
            fail("llms.txt has no blockquote summary")
        if not re.search(r"^## ", text, re.MULTILINE):
            fail("llms.txt has no link sections")
        if not re.search(r"^- \[", text, re.MULTILINE):
            fail("llms.txt has no link entries")

    full = SITE_DIR / "llms-full.txt"
    if not full.exists():
        fail("llms-full.txt is missing")
    elif full.stat().st_size < 20_000:
        fail(f"llms-full.txt is only {full.stat().st_size} bytes - suspiciously small")

    # Every built page needs a Markdown mirror beside it.
    for page in html_files:
        mirror = page.with_suffix(".md")
        if not mirror.exists():
            fail(f"{page_name(page)}: no Markdown mirror alongside it")
            continue
        body = mirror.read_text(encoding="utf-8")
        if not body.strip():
            fail(f"{page_name(mirror)}: mirror is empty")
        if "--8<--" in body:
            fail(f"{page_name(mirror)}: mirror still contains unresolved snippets")
        if body.lstrip().startswith("---"):
            fail(f"{page_name(mirror)}: mirror still contains front matter")


def audit_robots() -> None:
    robots = SITE_DIR / "robots.txt"
    if not robots.exists():
        fail("robots.txt is missing")
        return
    text = robots.read_text(encoding="utf-8")
    for agent in REQUIRED_AI_AGENTS:
        if not re.search(rf"^User-agent:\s*{re.escape(agent)}\s*$", text, re.M | re.I):
            fail(f"robots.txt has no explicit policy for {agent}")
    if "Sitemap:" not in text:
        fail("robots.txt does not advertise the sitemap")


def audit_faq() -> None:
    faq = SITE_DIR / "faq" / "index.html"
    if not faq.exists():
        fail("FAQ page is missing")
        return
    html = faq.read_text(encoding="utf-8", errors="replace")
    blocks = re.findall(
        r'<script type=["\']?application/ld\+json["\']?>(.*?)</script>', html, re.DOTALL
    )
    parsed = []
    for block in blocks:
        try:
            parsed.append(json.loads(block))
        except json.JSONDecodeError as exc:
            fail(f"faq/index.html: invalid JSON-LD block: {exc}")
    faq_blocks = [b for b in parsed if b.get("@type") == "FAQPage"]
    if not faq_blocks:
        fail("faq/index.html: no FAQPage structured data")
    elif len(faq_blocks[0].get("mainEntity", [])) < 5:
        fail("faq/index.html: FAQPage has fewer than 5 questions - parser regression?")

    troubleshooting = SITE_DIR / "troubleshooting" / "index.html"
    if not troubleshooting.exists():
        fail("troubleshooting page is missing")


def audit_jsonld(html_files: list[Path]) -> None:
    for path in html_files:
        html = path.read_text(encoding="utf-8", errors="replace")
        for block in re.findall(
            r'<script type=["\']?application/ld\+json["\']?>(.*?)</script>',
            html,
            re.DOTALL,
        ):
            try:
                json.loads(block)
            except json.JSONDecodeError as exc:
                fail(f"{page_name(path)}: invalid JSON-LD: {exc}")


def main() -> int:
    if not SITE_DIR.is_dir():
        print(f"ERROR: {SITE_DIR} not found - run `mkdocs build` first.")
        return 1

    html_files = sorted(
        p for p in SITE_DIR.rglob("*.html") if p.name != "404.html"
    )
    if not html_files:
        print("ERROR: no HTML pages found in the built site.")
        return 1

    for path in html_files:
        audit_html(path)
    audit_internal_links(html_files)
    audit_jsonld(html_files)
    audit_sitemap()
    audit_llms(html_files)
    audit_robots()
    audit_faq()

    if failures:
        print(f"Site audit FAILED - {len(failures)} issue(s) across {checked} pages:")
        for message in failures:
            print(f"  - {message}")
        return 1

    print(f"Site audit passed - {checked} pages, 0 issues.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
