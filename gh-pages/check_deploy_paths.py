#!/usr/bin/env python3
"""Verify that every source mirrored by hooks.py triggers a Pages rebuild."""

from __future__ import annotations

import fnmatch
import re
import sys
from pathlib import Path

import yaml

GH_PAGES = Path(__file__).resolve().parent
REPO_ROOT = GH_PAGES.parent
WORKFLOW = REPO_ROOT / ".github" / "workflows" / "deploy-gh-pages.yml"
READ_CALL = re.compile(r'_read\(\s*"([^"]+)"')

sys.path.insert(0, str(GH_PAGES))
from hooks import SKILL_SOURCES  # noqa: E402


def main() -> int:
    sources = set(READ_CALL.findall((GH_PAGES / "hooks.py").read_text(encoding="utf-8")))
    sources.update(f"skills/shared/{name}" for name in SKILL_SOURCES)
    config = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
    triggers = config.get("on", config.get(True, {})) or {}
    patterns = list((triggers.get("push") or {}).get("paths") or [])

    missing = [
        source
        for source in sorted(sources)
        if not (REPO_ROOT / source).is_file()
        or not any(fnmatch.fnmatch(source, pattern) for pattern in patterns)
    ]
    if missing:
        print("Deploy path coverage failed:", file=sys.stderr)
        for source in missing:
            print(f"  - {source}", file=sys.stderr)
        return 1

    print(f"Deploy path coverage passed for {len(sources)} mirrored sources.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
