---
"powerpointmcp": patch
---

Aligned the project's public-facing docs with the leading sister project
`mcp-server-excel`: ported the daily GitHub star-history chart. Added
`scripts/Update-StarHistory.ps1` (generates a theme-aware SVG from live
stargazer data) and wired it into the `deploy-gh-pages` workflow with a daily
schedule, then added a "GitHub Star History" section to both the README and the
docs-site homepage. The generated SVG is produced in CI and gitignored, mirroring
Excel's setup. Repository description, homepage URL, and topics were also updated
to match Excel's format.
