---
title: Frequently Asked Questions
description: Answers to common questions about PowerPoint MCP Server — Windows requirements, supported PowerPoint versions, AI client compatibility, security, and how it differs from python-pptx.
keywords: "PowerPoint MCP FAQ, PowerPoint MCP requirements, PowerPoint automation questions, does PowerPoint MCP work with Claude, PowerPoint MCP Windows only, python-pptx alternative"
---

# Frequently Asked Questions

<!--
  Every `###` heading on this page is parsed by gh-pages/hooks.py into FAQPage
  JSON-LD (see `_faq_jsonld`). The structured data is derived from this visible
  content, so there is nothing to keep in sync by hand — just write the
  questions here. Headings (rather than collapsible admonitions) are used
  deliberately: they give each answer a stable anchor that can be deep-linked
  from other pages and from search results.
-->

## Requirements and platform

### Does PowerPoint MCP Server run on macOS or Linux?

No. It drives a live Microsoft PowerPoint desktop process through Windows COM
interop, which exists only on Windows. There is no cross-platform or headless
mode, and none is planned. If you need cross-platform `.pptx` generation
without PowerPoint, an offline library such as `python-pptx` is the right tool —
see the [comparison page](comparison.md).

### Which versions of Microsoft PowerPoint are supported?

Microsoft PowerPoint **2016 or later** (desktop), including Microsoft 365
desktop builds.

Not supported, because they do not expose the PowerPoint COM automation
interface:

- PowerPoint Online (the browser version)
- PowerPoint Mobile
- LibreOffice Impress, Google Slides, Keynote

### Do I need a Microsoft 365 subscription?

No. Any licensed, **activated** desktop installation of PowerPoint 2016 or
later works, including perpetual-license editions. Activation matters:
PowerPoint refuses automation requests while it is in an unactivated or
reduced-functionality state.

### Can I run it on a server or in CI?

No. It requires an interactive Windows desktop session with PowerPoint
installed. Microsoft explicitly does not support Office automation in
non-interactive server contexts, and this project follows that guidance. Use it
on a workstation, not a build agent.

## Clients and setup

### Which AI assistants and MCP clients work with it?

Any MCP-compatible client that speaks stdio. There are step-by-step guides for
the most common ones:

- [Claude Desktop](claude-desktop.md)
- [VS Code and GitHub Copilot](vs-code.md)
- [Cursor](cursor.md)

Claude Code, Windsurf, Cline and Continue work the same way — point them at the
`mcp-powerpoint` command.

### Should I use the MCP Server or the CLI?

| Use | When |
|---|---|
| **CLI** (`pptcli`) | Coding agents and scripting. One command, no large tool schemas, so it is cheaper at high throughput. |
| **MCP Server** (`mcp-powerpoint`) | Conversational assistants that benefit from rich tool discovery and a persistent session. |

Both are first-class entry points built from the same core, so every operation
behaves identically on either surface.

### Do I need to close the presentation before automating it?

Yes. Close any open instance of the target file first — the server needs
exclusive access to the presentation while it drives PowerPoint.

## Behaviour

### Does it edit `.pptx` files directly, or does it open PowerPoint?

It opens and drives the **real PowerPoint application**. PowerPoint itself
renders and writes the file, so the output is always a file PowerPoint can
open, with true-fidelity theme, layout and font rendering. Offline `.pptx`
libraries write the XML themselves and can produce files that render
differently or fail to open.

### What is export-to-verify?

After any visual edit, the AI exports the slide — or the whole deck — to images
and inspects the rendered result:

```text
export(action="export-slide-to-image", session_id="...", slide_index=3, output_path="C:\\out\\slide3.png")
```

A vision-capable model can then catch overlapping shapes, text overflow and
layout regressions that text-only automation simply cannot detect. This is the
project's core differentiator.

### Are slide and shape indexes 0-based or 1-based?

**1-based, everywhere.** Slide 1 is the first slide, matching PowerPoint's
native COM object model. Table rows and columns are 1-based too.

### Why does `POWERPNT.EXE` stay running for a while after closing a session?

That is normal Office behaviour. PowerPoint's own post-quit cleanup can take up
to roughly 90–200 seconds after a session is closed. The session is removed
from the registry immediately and the underlying process is disposed on a
background task, so you never have to wait for it. See
[troubleshooting](troubleshooting.md#lingering-powerpntexe-processes).

## Comparison

### How is this different from `python-pptx`?

`python-pptx` manipulates the Office Open XML inside a `.pptx` file without
PowerPoint running. That makes it cross-platform and CI-friendly, but it cannot
render, cannot resolve theme and layout inheritance the way PowerPoint does, and
cannot show an AI what a slide actually looks like.

PowerPoint MCP Server drives real PowerPoint on Windows and can export any slide
to an image for visual verification. Full breakdown on the
[comparison page](comparison.md).

## Privacy and licensing

### Does it send my presentation content anywhere?

No. The server runs entirely on your machine, performs local COM calls,
collects no telemetry and does not phone home. Only your own AI client sends
content to whichever model provider you have configured. See the
[privacy policy](privacy.md).

### Is it free and open source?

Yes — MIT licensed and developed in the open
[on GitHub](https://github.com/sbroenne/mcp-server-powerpoint).

## Still stuck?

- [Troubleshooting guide](troubleshooting.md) — concrete errors and fixes
- [Installation guide](installation.md) — every install path
- [Open an issue](https://github.com/sbroenne/mcp-server-powerpoint/issues)
