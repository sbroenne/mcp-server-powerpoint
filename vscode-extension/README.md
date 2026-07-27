# PowerPoint MCP Server — VS Code Extension

Automate Microsoft PowerPoint from AI assistants (GitHub Copilot and other MCP
clients) directly inside VS Code. The extension registers the
[PowerPoint MCP Server](https://powerpointmcpserver.dev) via VS Code's MCP
server definition provider and ships the associated Agent Skills.

> **Windows only.** Requires Microsoft PowerPoint (desktop) to be installed.

## What it does

- Registers the `powerpoint-mcp` MCP server so Copilot Chat can create and edit
  PowerPoint decks — slides, shapes, text, tables, charts, images, notes,
  layouts — and export slides to images for visual verification.
- Bundles a **self-contained** MCP server executable (no .NET runtime needed).
- Contributes the PowerPoint Agent Skill (`chatSkills`) for better guidance.

## Getting started

1. Install the extension from the VS Code Marketplace.
2. Ensure Microsoft PowerPoint is installed.
3. Open Copilot Chat and ask it to build or edit a `.pptx` deck.

## Building locally

```pwsh
cd vscode-extension
npm install
npm run compile        # quick check: compiles src/extension.ts -> out/extension.js only
npm run package        # produce the .vsix (runs vscode:prepublish -> full bundle)
```

`npm run compile` is just a fast TypeScript build for local iteration — it does
**not** bundle the MCP server or skills. The `.vsix` is produced by
`npm run package`, which triggers `vscode:prepublish`: this publishes the
self-contained MCP server into `bin/`, copies the skill pack and changelog, and
then compiles the TypeScript.

## Links

- Documentation: https://powerpointmcpserver.dev
- Source: https://github.com/sbroenne/mcp-server-powerpoint

Licensed under the MIT License.
