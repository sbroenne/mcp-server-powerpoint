---
"powerpointmcp": patch
---

**Better docs and discoverability**: new task guides, a full reference section,
FAQ, troubleshooting, comparison, and per-client setup guides (Claude Desktop,
VS Code, Cursor) on [powerpointmcpserver.dev](https://powerpointmcpserver.dev/).

Five step-by-step [guides](https://powerpointmcpserver.dev/guides/) cover
building a deck with AI, editing an existing presentation, exporting slides to
images, automating charts and tables, and applying a corporate template. The
complete expert [reference](https://powerpointmcpserver.dev/reference/) that
ships inside the agent skill packages is now published on the site too, so the
website and the guidance an assistant actually receives can never disagree.

Every documentation page is now also available as plain Markdown — just append
`index.md` to any page URL — and the site publishes `llms.txt` and
`llms-full.txt`, so AI assistants can read the docs directly and quote them
accurately. AI crawlers are explicitly welcomed in `robots.txt`. Pages now carry
complete social-preview metadata and richer structured data, including an FAQ
that search engines can surface directly.

Also fixes stale guidance that claimed slides could not be reordered — they can,
with `slide(action: "move-to")` — and a stale count in the MCPB bundle manifest,
which advertised 137 operations instead of 141. The CLI agent skill's command
reference was 29 operations out of date, missing the entire SmartArt domain; it
is now generated from the same source as the rest of the skill and checked in
CI. NuGet and MCP registry listings now point at the documentation site, and the
README is shorter and links to the docs instead of duplicating them.
