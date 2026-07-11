---
"powerpointmcp": patch
---

Improved SEO for the public docs site (powerpointmcpserver.dev): added a
branded 1200x630 Open Graph/Twitter card image, richer `SoftwareApplication`
structured data (image, MIT license, `sameAs` links to GitHub and NuGet), a
`WebSite`+`Person` JSON-LD graph, and per-page `BreadcrumbList` structured
data for interior pages. Also sped up `scripts/pre-commit.ps1` for
docs-only commits (including `gh-pages/` website changes) by fully skipping
the Release build and Core test gates when no compiled code changed.
