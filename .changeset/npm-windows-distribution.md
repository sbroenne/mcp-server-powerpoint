---
"powerpointmcp": minor
---

Add npm launchers and native Windows x64/ARM64 runtime packages for the MCP server and `pptcli`. Agent plugins and the Claude Desktop bundle now launch the public npm packages through `npx`, so those installation paths require Node.js 18 or later.

Modernize MCP responses with structured content, generated output schemas, stricter pre-dispatch argument validation, and cancellation-aware calls while retaining text JSON compatibility.

Harden the metadata-only MCPB build, add native x64 and ARM64 VS Code packages, and improve extension startup diagnostics for Windows and desktop PowerPoint prerequisites.
