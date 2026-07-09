---
title: Privacy Policy
description: Privacy policy for PowerPoint MCP Server. How we handle your data, telemetry collection and data protection.
---

# Privacy Policy

## Summary

PowerPoint MCP Server runs entirely on your own machine. It does not collect
telemetry, does not phone home, and does not transmit your presentation
content anywhere. All automation happens locally through COM calls to a
PowerPoint process running on your Windows PC.

## What data is processed

- **Presentation files** you explicitly open, create, edit, or export are
  processed locally by PowerPoint itself, driven by MCP tool calls. Nothing
  is uploaded.
- **Tool call arguments and results** (file paths, slide indexes, text
  content, etc.) stay within the local MCP stdio session between your AI
  assistant and the server process — they are never sent to a third-party
  service by this project.
- **Exported images** (from `export_slide_to_image` /
  `export_all_slides_to_images`) are written only to the local file paths
  you specify.

## What this project does NOT do

- ❌ No analytics, telemetry, or usage tracking
- ❌ No network calls to servers operated by this project
- ❌ No collection of personal information
- ❌ No storage of your presentation content outside your own file system

## Third-party AI assistants

If you use this MCP server with a third-party AI assistant (GitHub Copilot,
Claude Desktop, Cursor, etc.), that assistant's own privacy policy governs
how it handles the conversation context it sends to its model provider.
This project only controls the local MCP server process and the PowerPoint
automation it performs.

## Questions

Open an issue on the
[GitHub repository](https://github.com/sbroenne/mcp-server-powerpoint) for
any privacy-related questions.
