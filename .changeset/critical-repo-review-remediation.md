---
"powerpointmcp": patch
---

**Critical repo review remediation** — a pass comparing this repo against
`mcp-server-excel` (the architectural template) surfaced several gaps, now
fixed:

- Added `get-font-size`, `get-bold`, and `get-font-color` operations to the
  `textframe` tool (previously only the `set-` variants existed), bringing
  the total to 137 operations across 13 tools.
- Fixed several COM reference leaks in the Chart/Shape/Image/Animation/
  Presentation/Slide commands — every manually-acquired COM object is now
  released via `ComUtilities.Release` in a `finally` block.
- Fixed ComInterop lifecycle bugs and removed dead code left over from the
  original Excel-to-PowerPoint port (`ServiceRegistryGenerator`'s unused
  `GetShortAlias` helper and stray `excelcli` string references).
- Reverted an in-progress window-hiding change that had broken embedded-chart
  OLE activation — PowerPoint windows remain visible, matching documented
  behavior.
- Aligned CI workflows, `Directory.Build.*`, `Packages.props`, manifest,
  `.gitattributes`, `.editorconfig`, `SECURITY.md`, `PRIVACY.md`, dependabot
  config, and issue templates with the `mcp-server-excel` template repo.
