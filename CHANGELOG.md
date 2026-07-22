# Changelog

## [0.1.1] - 2026-07-22

### Minor Changes

- [#32](https://github.com/sbroenne/mcp-server-powerpoint/pull/32) [`778240d`](https://github.com/sbroenne/mcp-server-powerpoint/commit/778240d92361e9a488847cef1502d0db5ae71ab7) Thanks [@sbroenne](https://github.com/sbroenne)! - Aligned mcp-server-powerpoint's architecture and tooling with mcp-server-excel, the
  authoritative template for this family of projects:

  - **Single MCP dispatch tool**: replaced the 12 separate presentation-related MCP tools
    with one `presentation` tool taking an `action` parameter, matching Excel's dispatch
    pattern. Added matching CLI session-management commands so the CLI and MCP Server
    remain equal, parity-checked entry points.
  - **COM object cleanup**: every `dynamic` COM object obtained inside a Core Commands
    method is now released in a `finally` block once no longer needed, instead of relying
    solely on the top-level session Application/Presentation cleanup. This closes a class
    of potential COM object leaks under heavy or long-running use.
  - **Crash-safety process tracking**: PowerPoint process IDs are now tracked in a
    process-wide registry with an `AppDomain.ProcessExit` handler that force-kills any
    still-running `POWERPNT.exe` if the host process itself terminates uncleanly (crash,
    forced kill, etc.), closing a gap versus Excel's existing crash-safety net.
  - **Ported audit scripts** from mcp-server-excel: `check-com-leaks.ps1`,
    `check-success-flag.ps1`, `check-dynamic-casts.ps1`, and a new
    `check-core-interface-completeness.ps1` tailored to this project's
    generate-enums-from-interface architecture. All are wired into the pre-commit hook.
  - Corrected stale documentation (tool/operation/domain counts) across the README,
    docs site, and skills files.

  No behavior change for existing MCP tool calls beyond the dispatch-tool consolidation;
  CLI commands are unaffected.

### Patch Changes

- [#26](https://github.com/sbroenne/mcp-server-powerpoint/pull/26) [`ddb6808`](https://github.com/sbroenne/mcp-server-powerpoint/commit/ddb68080460e6f1a6b9e8d0c86e309874cf6ae1d) Thanks [@sbroenne](https://github.com/sbroenne)! - Aligned the project's public-facing docs with the leading sister project
  `mcp-server-excel`: ported the daily GitHub star-history chart. Added
  `scripts/Update-StarHistory.ps1` (generates a theme-aware SVG from live
  stargazer data) and wired it into the `deploy-gh-pages` workflow with a daily
  schedule, then added a "GitHub Star History" section to both the README and the
  docs-site homepage. The generated SVG is produced in CI and gitignored, mirroring
  Excel's setup. Repository description, homepage URL, and topics were also updated
  to match Excel's format.

- [#33](https://github.com/sbroenne/mcp-server-powerpoint/pull/33) [`9663e7a`](https://github.com/sbroenne/mcp-server-powerpoint/commit/9663e7a7be95ff6778c22b21de5d833be801c351) Thanks [@sbroenne](https://github.com/sbroenne)! - **Fixed stale tool/operation counts** (#33): the docs advertised "132
  operations" even though the `image` tool gained `set-crop`/`get-crop` a while
  back, bringing the real total to 134 operations across 13 tools. Fixed across
  the README, MCP Server README, mcpb README, and the public docs site.

  Also replaced the old path-filtered build workflows with a single CI Gate
  workflow that runs on every pull request, and added a documentation
  count-check to the pre-commit hook so this kind of drift is caught
  automatically going forward.

- [#35](https://github.com/sbroenne/mcp-server-powerpoint/pull/35) [`3f72dae`](https://github.com/sbroenne/mcp-server-powerpoint/commit/3f72dae2d4203af5233fe90245e3e0bc8772ed41) Thanks [@sbroenne](https://github.com/sbroenne)! - **Critical repo review remediation** — a pass comparing this repo against
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

- [#29](https://github.com/sbroenne/mcp-server-powerpoint/pull/29) [`8a268ca`](https://github.com/sbroenne/mcp-server-powerpoint/commit/8a268ca1bfd2adfc57a59a0b39c87ed8f98c089f) Thanks [@sbroenne](https://github.com/sbroenne)! - Improved SEO for the public docs site (powerpointmcpserver.dev): added a
  branded 1200x630 Open Graph/Twitter card image, richer `SoftwareApplication`
  structured data (image, MIT license, `sameAs` links to GitHub and NuGet), a
  `WebSite`+`Person` JSON-LD graph, and per-page `BreadcrumbList` structured
  data for interior pages. Also sped up `scripts/pre-commit.ps1` for
  docs-only commits (including `gh-pages/` website changes) by fully skipping
  the Release build and Core test gates when no compiled code changed.

## [0.1.0] - 2026-07-11

### Minor Changes

- [#23](https://github.com/sbroenne/mcp-server-powerpoint/pull/23) [`5b5b21a`](https://github.com/sbroenne/mcp-server-powerpoint/commit/5b5b21a77cff2f7dd666d1bf404bc0ac781f9bc2) Thanks [@sbroenne](https://github.com/sbroenne)! - Add `chart(action: "replace-chart-data", ...)` to wholesale-replace an existing chart's
  categories, series names, and values in a single call, including changing the category count.
  `series_values` is a flat, series-major array (all values for the first series, then the second,
  etc.), avoiding the previous delete-shape-and-recreate workaround for editing chart data in place.

- [#19](https://github.com/sbroenne/mcp-server-powerpoint/pull/19) [`e8b989d`](https://github.com/sbroenne/mcp-server-powerpoint/commit/e8b989d48180812f905000e21810ee9aa33b7792) Thanks [@sbroenne](https://github.com/sbroenne)! - Add document properties management to the Presentation domain: `set_document_property` /
  `get_document_property` for built-in metadata (Title, Subject, Author, Keywords, Comments,
  Category, Manager, Company), and `set_custom_property` / `get_custom_property` /
  `remove_custom_property` for user-defined custom document properties.

- [#22](https://github.com/sbroenne/mcp-server-powerpoint/pull/22) [`4db42d8`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4db42d843f7f446b5515553d4b868c653265a524) Thanks [@sbroenne](https://github.com/sbroenne)! - Add gradient background support to the `slide` and `master` domains:
  `set-gradient-background` / `get-gradient-background` set or read a two-color gradient fill
  (style is an `MsoGradientStyle` name — `msoGradientHorizontal`, `msoGradientVertical`,
  `msoGradientDiagonalUp`, `msoGradientDiagonalDown`, `msoGradientFromCorner`,
  `msoGradientFromTitle`, `msoGradientFromCenter` — plus a `1`-`4` variant). `slide`'s version
  overrides a single slide's background (same override semantics as the existing
  `set-background-color`); `master`'s version sets the gradient for every slide that inherits from
  the master.

- [#17](https://github.com/sbroenne/mcp-server-powerpoint/pull/17) [`380b289`](https://github.com/sbroenne/mcp-server-powerpoint/commit/380b28930f57510e08be41dd341d67447e516d36) Thanks [@sbroenne](https://github.com/sbroenne)! - Add hyperlink management to the Shape domain: `set-hyperlink`, `get-hyperlink`, and `remove-hyperlink` actions let callers attach, inspect, and clear mouse-click hyperlinks (with optional screen tip text) on any shape.

- [#21](https://github.com/sbroenne/mcp-server-powerpoint/pull/21) [`4c8f7c4`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4c8f7c49dc92de6eaf723968b00dc745d34272b3) Thanks [@sbroenne](https://github.com/sbroenne)! - Add picture-effect actions to the `image` domain: `set-brightness-contrast`/`get-brightness-contrast`
  (adjust brightness and contrast, each a float in `[0, 1]`) and `set-recolor`/`get-recolor` (recolor
  a picture to grayscale, black-and-white, watermark, or automatic/no recolor via
  `msoPictureAutomatic`, `msoPictureGrayscale`, `msoPictureBlackAndWhite`, `msoPictureWatermark`).

- [#20](https://github.com/sbroenne/mcp-server-powerpoint/pull/20) [`845737e`](https://github.com/sbroenne/mcp-server-powerpoint/commit/845737edf9fb30d8bfcada25555bd90e243fe2ed) Thanks [@sbroenne](https://github.com/sbroenne)! - Add parameterized shape effects: drop shadow (color/transparency/blur/offset), glow, reflection, soft edge, and 3D bevel. Extends the existing `set-shadow`/`get-shadow` actions with optional formatting parameters (backward compatible) and adds new `set-glow`/`get-glow`, `set-reflection`/`get-reflection`, `set-soft-edge`/`get-soft-edge`, and `set-bevel`/`get-bevel` actions to the `shape` MCP tool.

- [#15](https://github.com/sbroenne/mcp-server-powerpoint/pull/15) [`41dd351`](https://github.com/sbroenne/mcp-server-powerpoint/commit/41dd35106e30a7ec66c87ba747eb206bd0e58179) Thanks [@sbroenne](https://github.com/sbroenne)! - Added a new `smartart` tool for creating and editing SmartArt diagrams: `add-smart-art` (from
  any built-in gallery layout, e.g. "Basic Process", "Organization Chart"), `add-node`,
  `add-child-node`, `set-node-text`, `get-node-text`, `delete-node`, and `get-node-count`.

- [#24](https://github.com/sbroenne/mcp-server-powerpoint/pull/24) [`a2fc594`](https://github.com/sbroenne/mcp-server-powerpoint/commit/a2fc59460e83379c12743982b4ca18efeb47ba53) Thanks [@sbroenne](https://github.com/sbroenne)! - Add `textframe(action: "set-autosize"/"get-autosize", ...)` to control a shape's text-frame
  auto-fit behavior (`ppAutoSizeNone`, `ppAutoSizeShapeToFitText`, `ppAutoSizeTextToFitShape`),
  matching PowerPoint's "Resize shape to fit text" / "Shrink text on overflow" options.

### Patch Changes

- [#16](https://github.com/sbroenne/mcp-server-powerpoint/pull/16) [`fbaf960`](https://github.com/sbroenne/mcp-server-powerpoint/commit/fbaf960051571bd6ca92a0d37db5dfe93c85947f) Thanks [@sbroenne](https://github.com/sbroenne)! - Add a short "Also automating spreadsheets?" tip right after the README's hero section, linking to Excel MCP Server, mirroring the same repositioning done on the docs homepage.

- [#12](https://github.com/sbroenne/mcp-server-powerpoint/pull/12) [`fc05b46`](https://github.com/sbroenne/mcp-server-powerpoint/commit/fc05b46d153083f3e437f779a5d48b45dab3fffc) Thanks [@sbroenne](https://github.com/sbroenne)! - Restructured the documentation site homepage to match the sister
  `mcp-server-excel` project's structure: removed the redundant "Quick install"
  card grid from the homepage (installation already has its own dedicated
  page), split "How it works" out into a new `architecture.md` page, and split
  "Related projects" out into a new `related-projects.md` page. Also fixed the
  last remaining stale "31 tools across 10 domains" references in
  `installation.md`, `mcp-server.md`, and `mcpb/README.md` to the current
  18-tool / ~98-operation / 12-domain count.

- [#25](https://github.com/sbroenne/mcp-server-powerpoint/pull/25) [`0f4fcd8`](https://github.com/sbroenne/mcp-server-powerpoint/commit/0f4fcd8da984fb8c61ec64ecaa6c394d3af1d6bb) Thanks [@sbroenne](https://github.com/sbroenne)! - **Release automation: auto-merge the changelog PR.** The post-release step that opens the `chore/changelog-vX` PR now also merges it (queued auto-merge, falling back to an immediate squash merge). Previously the PR was only created and left open until a maintainer merged it by hand, which could leave releases with a stale/missing CHANGELOG on `main`.

- [#11](https://github.com/sbroenne/mcp-server-powerpoint/pull/11) [`87e5d42`](https://github.com/sbroenne/mcp-server-powerpoint/commit/87e5d421daa2b1577010da47b4ca9848ea1128cf) Thanks [@sbroenne](https://github.com/sbroenne)! - Fixed the "What You Can Do" section of the root README and rewrote
  `gh-pages/docs/features.md` and `src/PowerPointMcp.McpServer/README.md`, which
  all still described the old 31-tool/10-domain architecture. They now
  correctly document the current 18-tool, ~98-operation, 12-domain surface,
  including the previously undocumented Template, Master, and Animation
  domains.

- [#10](https://github.com/sbroenne/mcp-server-powerpoint/pull/10) [`9b0b70e`](https://github.com/sbroenne/mcp-server-powerpoint/commit/9b0b70e0ee470ff63a40f1df8f9facd574ab4c6e) Thanks [@sbroenne](https://github.com/sbroenne)! - Fixed the root README: corrected the "How It Works" diagram to reflect the
  actual shared `PowerPointMcp Service` + Core Commands architecture (it
  previously omitted the service layer), and removed a broken/non-existent
  "HeyGen MCP Server" link from Related Projects.

- [#27](https://github.com/sbroenne/mcp-server-powerpoint/pull/27) [`794366c`](https://github.com/sbroenne/mcp-server-powerpoint/commit/794366c0e3fcd351c842f69544aea414ae258845) Thanks [@sbroenne](https://github.com/sbroenne)! - Fix the release workflow's "Update Extension Version" step failing with "Version not changed"
  whenever the release version coincidentally matches the VS Code extension's fixed baseline
  version (0.1.0, never committed back to the repo since the bump is ephemeral per-run). Adds
  `--allow-same-version` to the `npm version` call.

- [#14](https://github.com/sbroenne/mcp-server-powerpoint/pull/14) [`94a7cf0`](https://github.com/sbroenne/mcp-server-powerpoint/commit/94a7cf0215e1ff7246ec7f538f79c54d13df41b3) Thanks [@sbroenne](https://github.com/sbroenne)! - Move the "Also automating spreadsheets?" sister-project tip on the docs homepage to appear directly under the hero callout instead of at the bottom of the page.

## [0.0.2] - 2026-07-09

### Patch Changes

- [#7](https://github.com/sbroenne/mcp-server-powerpoint/pull/7) [`4bd5377`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4bd5377c359017f75a7af74c0036aa76e2396a53) Thanks [@sbroenne](https://github.com/sbroenne)! - Replaced manual `CHANGELOG.md` editing and the fragile awk/sed-based release-notes
  extraction with [changesets](https://github.com/changesets/changesets): contributors
  now add a small `.changeset/*.md` fragment describing their change, CI enforces one
  is present (or the `skip-changelog` label), and `scripts/Build-Changelog.ps1`
  compiles pending fragments into `CHANGELOG.md` and the GitHub Release body at
  release time. See `docs/RELEASE-STRATEGY.md` for the full process.

- [#7](https://github.com/sbroenne/mcp-server-powerpoint/pull/7) [`4bd5377`](https://github.com/sbroenne/mcp-server-powerpoint/commit/4bd5377c359017f75a7af74c0036aa76e2396a53) Thanks [@sbroenne](https://github.com/sbroenne)! - Added PowerPoint-branded icon assets for the VS Code extension, the MCPB bundle,
  and the gh-pages documentation site, replacing the generic/missing icons used
  previously.

All notable changes to the PowerPoint MCP Server are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.1] - 2026-07-09

### Added

- **MCP server** with 31 tools across 10 domains: Presentation (create, open,
  save, close, list sessions), Slide (add, count, delete), Shape (rectangle, text
  box, count, delete, position, size), Text (set/get text, font size, bold, font
  color), Table (add, set/get cell text), Notes (set/get), Layout (set/get),
  Image (add picture), Chart (add, get data), and Export (slide / all slides to
  image) for visual verification.
- **Live PowerPoint COM automation** via an STA thread with an OLE message
  filter; one long-lived session per open presentation, addressed by `sessionId`.
- **Resilient shutdown**: sessions are disposed on host shutdown with exponential
  backoff so no `POWERPNT.exe` process is orphaned.
- **Non-blocking `create_presentation`**: creates and keeps the deck open,
  returning a `sessionId` immediately (~2s instead of ~90-210s).
- **Distribution**: NuGet .NET tools (`Sbroenne.PowerPointMcp.McpServer` →
  `mcp-powerpoint`, `Sbroenne.PowerPointMcp.CLI` → `pptcli`), MCP Registry
  manifest, and a Claude Desktop MCPB bundle.
- **Agent skill pack** and documentation site at
  [powerpointmcpserver.dev](https://powerpointmcpserver.dev).

[0.0.1]: https://github.com/sbroenne/mcp-server-powerpoint/releases/tag/v0.0.1
