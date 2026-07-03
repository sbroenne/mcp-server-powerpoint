---
layout: default
title: "Security Policy"
description: "Security policy for PowerPoint MCP Server. How to report vulnerabilities, supported versions, and security features."
permalink: /security/
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <h1 class="hero-title">Security Policy</h1>
      <p class="hero-subtitle">How to report vulnerabilities, supported versions, and security features</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">

## Supported Versions

PowerPoint MCP Server is currently in early development (pre-1.0). Only the latest published
release receives security fixes until a formal support matrix is established after the first
stable release.

## Reporting a Vulnerability

Please report suspected security vulnerabilities privately using
[GitHub's private vulnerability reporting](https://github.com/sbroenne/mcp-server-powerpoint/security/advisories/new)
for this repository, rather than opening a public issue. Include:

- A description of the vulnerability and its potential impact
- Steps to reproduce (a minimal PowerPoint file/tool-call sequence, if applicable)
- Any relevant logs (redact file paths or content you consider sensitive)

We aim to acknowledge reports within a few business days.

## Security Model

- **Local, Windows-only automation.** The MCP server drives a real, local PowerPoint desktop
  process over COM (`Microsoft.Office.Interop.PowerPoint`). It does not open network ports, does
  not send presentation content anywhere, and only operates on files you explicitly open, create,
  or export.
- **Native Office API only.** Because it uses PowerPoint's own COM automation surface rather than
  a third-party `.pptx` parser, there is no custom file-format parsing that could be exploited by
  a malformed file — PowerPoint itself validates and opens the file.
- **No macros executed by default.** Tools do not execute arbitrary VBA or external code; they
  call discrete, typed automation operations (add shape, set text, export image, etc.).
- **Export-to-verify.** Visual verification via `export_slide_to_image` / `export_all_slides_to_images`
  writes image files only to paths you specify.

See also the [Privacy Policy](/privacy/) for how data is (not) collected or transmitted.

</div>
