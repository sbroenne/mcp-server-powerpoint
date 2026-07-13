---
"powerpointmcp": patch
---

**Fixed stale tool/operation counts** (#33): the docs advertised "132
operations" even though the `image` tool gained `set-crop`/`get-crop` a while
back, bringing the real total to 134 operations across 13 tools. Fixed across
the README, MCP Server README, mcpb README, and the public docs site.

Also replaced the old path-filtered build workflows with a single CI Gate
workflow that runs on every pull request, and added a documentation
count-check to the pre-commit hook so this kind of drift is caught
automatically going forward.
