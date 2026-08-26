#!/usr/bin/env python3
"""Audit built documentation for structure, metadata, accessibility, and discovery."""

from __future__ import annotations

import json
import re
import sys
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit

SITE = Path(__file__).resolve().parent / "_site"
FEATURES = Path(__file__).resolve().parent / "docs" / "features.md"
SITE_URL = "https://powerpointmcpserver.dev/"
AI_CRAWLERS = ("GPTBot", "OAI-SearchBot", "ClaudeBot", "PerplexityBot", "Google-Extended")
BADGE_HOSTS = ("img.shields.io", "vsmarketplacebadges.dev", "badgen.net")


class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.targets: list[str] = []
        self.images: list[dict[str, str | None]] = []
        self.h1_count = 0
        self.jsonld: list[str] = []
        self._jsonld_parts: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        target = values.get("href") if tag == "a" else values.get("src")
        if target:
            self.targets.append(target)
        if tag == "img":
            self.images.append(values)
        elif tag == "h1":
            self.h1_count += 1
        elif tag == "script" and values.get("type") == "application/ld+json":
            self._jsonld_parts = []

    def handle_endtag(self, tag: str) -> None:
        if tag == "script" and self._jsonld_parts is not None:
            self.jsonld.append("".join(self._jsonld_parts))
            self._jsonld_parts = None

    def handle_data(self, data: str) -> None:
        if self._jsonld_parts is not None:
            self._jsonld_parts.append(data)


def page_name(path: Path) -> str:
    return path.relative_to(SITE).as_posix()


def resolve_target(page: Path, raw_target: str) -> Path | None:
    parsed = urlsplit(raw_target)
    if parsed.scheme or parsed.netloc or raw_target.startswith(("mailto:", "tel:", "data:", "#")):
        return None

    path = unquote(parsed.path)
    if not path:
        return None
    target = SITE / path.lstrip("/") if path.startswith("/") else page.parent / path
    if path.endswith("/"):
        target /= "index.html"
    elif not target.suffix:
        target /= "index.html"
    return target.resolve()


def audit_page(path: Path, failures: list[str]) -> None:
    html = path.read_text(encoding="utf-8", errors="replace")
    name = page_name(path)
    parser = PageParser()
    parser.feed(html)

    canonical = re.search(r'<link[^>]+rel=["\']?canonical["\']?[^>]*>', html)
    if canonical is None or SITE_URL not in canonical.group(0):
        failures.append(f"{name}: missing canonical URL for {SITE_URL}")

    description = re.search(
        r'<meta[^>]+name=["\']?description["\']?[^>]+content=["\']([^"\']+)',
        html,
    )
    if description is None or not 50 <= len(" ".join(description.group(1).split())) <= 200:
        failures.append(f"{name}: missing or unsuitable meta description")

    for key in ("og:title", "og:description", "og:image", "og:url", "og:type"):
        if f'property="{key}"' not in html and f"property={key}" not in html:
            failures.append(f"{name}: missing {key}")
    for key in ("twitter:card", "twitter:title", "twitter:description", "twitter:image"):
        if f'name="{key}"' not in html and f"name={key}" not in html:
            failures.append(f"{name}: missing {key}")

    if parser.h1_count != 1:
        failures.append(f"{name}: found {parser.h1_count} h1 elements; expected one")

    if 'type="text/markdown"' not in html and "type=text/markdown" not in html:
        failures.append(f"{name}: missing Markdown alternate link")
    mirror = path.with_suffix(".md")
    if not mirror.is_file():
        failures.append(f"{name}: missing Markdown mirror")
    else:
        mirror_text = mirror.read_text(encoding="utf-8")
        if not mirror_text.strip() or "8<--" in mirror_text or mirror_text.startswith("---"):
            failures.append(f"{page_name(mirror)}: empty, unresolved, or contains front matter")

    if path.name != "index.html" or path.parent != SITE:
        jsonld_types = []
        for block in parser.jsonld:
            try:
                payload = json.loads(block)
            except json.JSONDecodeError as exc:
                failures.append(f"{name}: invalid JSON-LD: {exc}")
                continue
            jsonld_types.append(payload.get("@type"))
        if "BreadcrumbList" not in jsonld_types:
            failures.append(f"{name}: missing BreadcrumbList JSON-LD")
    else:
        for block in parser.jsonld:
            try:
                json.loads(block)
            except json.JSONDecodeError as exc:
                failures.append(f"{name}: invalid JSON-LD: {exc}")

    for image in parser.images:
        source = image.get("src") or ""
        parsed = urlsplit(source)
        exempt = parsed.netloc in BADGE_HOSTS or parsed.path.endswith(".svg")
        if not exempt and ("width" not in image or "height" not in image):
            failures.append(f"{name}: image lacks width/height: {source}")

    site_root = SITE.resolve()
    for raw_target in parser.targets:
        target = resolve_target(path, raw_target)
        if target is None:
            continue
        if site_root not in target.parents and target != site_root:
            failures.append(f"{name}: target escapes site: {raw_target}")
        elif not target.exists():
            failures.append(f"{name}: broken target: {raw_target}")

    search = re.search(r"<div[^>]*\bclass=[\"']?md-search[\"'\s>][^>]*>", html)
    if search is None or "aria-label" not in search.group(0):
        failures.append(f"{name}: search dialog lacks an accessible name")
    progress = re.search(r"<div[^>]*\bclass=[\"']?md-progress[\"'\s>][^>]*>", html)
    if progress is None or "aria-label" not in progress.group(0):
        failures.append(f"{name}: progress indicator lacks an accessible name")


def audit_sitemap(html_pages: list[Path], failures: list[str]) -> None:
    path = SITE / "sitemap.xml"
    if not path.is_file():
        failures.append("missing sitemap.xml")
        return
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as exc:
        failures.append(f"sitemap.xml is invalid: {exc}")
        return
    namespace = {"s": "http://www.sitemaps.org/schemas/sitemap/0.9"}
    urls = {
        node.text.rstrip("/") + "/"
        for node in root.findall("s:url/s:loc", namespace)
        if node.text
    }
    expected = {
        SITE_URL
        if page == SITE / "index.html"
        else SITE_URL + page.parent.relative_to(SITE).as_posix().strip("/") + "/"
        for page in html_pages
    }
    missing = sorted(expected - urls)
    if missing:
        failures.append(f"sitemap.xml is missing {len(missing)} built page URL(s)")
    if not (SITE / "sitemap.xml.gz").is_file():
        failures.append("missing sitemap.xml.gz")


def audit_discovery(html_pages: list[Path], failures: list[str]) -> None:
    llms = SITE / "llms.txt"
    full = SITE / "llms-full.txt"
    if not llms.is_file():
        failures.append("missing llms.txt")
    else:
        text = llms.read_text(encoding="utf-8")
        if not text.startswith("# ") or "> " not in text[:1000] or text.count("](") < len(html_pages):
            failures.append("llms.txt is incomplete")
    if not full.is_file():
        failures.append("missing llms-full.txt")
    else:
        text = full.read_text(encoding="utf-8")
        if "8<--" in text or len(text) < 20_000:
            failures.append("llms-full.txt is incomplete or contains unresolved snippets")

    tools_path = SITE / "tools.json"
    if not tools_path.is_file():
        failures.append("missing tools.json")
        return
    try:
        tools = json.loads(tools_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        failures.append(f"tools.json is invalid: {exc}")
        return
    feature_text = FEATURES.read_text(encoding="utf-8")
    headline = re.search(
        r"exposes \*\*(?P<tools>\d+) MCP tools with (?P<operations>\d+) operations",
        feature_text,
    )
    if headline is None:
        failures.append("cannot read feature totals")
        return
    expected_tools = int(headline.group("tools"))
    expected_operations = int(headline.group("operations"))
    if tools.get("toolCount") != expected_tools or tools.get("operationCount") != expected_operations:
        failures.append("tools.json totals do not match the feature reference")
    categories = tools.get("categories")
    if not isinstance(categories, list) or len(categories) != expected_tools:
        failures.append("tools.json categories do not enumerate every tool")


def audit_robots(failures: list[str]) -> None:
    path = SITE / "robots.txt"
    if not path.is_file():
        failures.append("missing robots.txt")
        return
    text = path.read_text(encoding="utf-8")
    for crawler in AI_CRAWLERS:
        if f"User-agent: {crawler}" not in text:
            failures.append(f"robots.txt has no policy for {crawler}")
    if f"Sitemap: {SITE_URL}sitemap.xml" not in text:
        failures.append("robots.txt does not declare the canonical sitemap")


def main() -> int:
    if not SITE.is_dir():
        print(f"ERROR: {SITE} not found; run mkdocs build first", file=sys.stderr)
        return 2

    failures: list[str] = []
    html_pages = sorted(
        path
        for path in SITE.rglob("*.html")
        if path.name != "404.html" and "assets" not in path.relative_to(SITE).parts
    )
    if not html_pages:
        failures.append("no HTML pages were built")

    for page in html_pages:
        audit_page(page, failures)
    audit_sitemap(html_pages, failures)
    audit_discovery(html_pages, failures)
    audit_robots(failures)

    if failures:
        print(f"Site audit failed with {len(failures)} issue(s):", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure}", file=sys.stderr)
        return 1

    print(f"Site audit passed for {len(html_pages)} HTML pages.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
