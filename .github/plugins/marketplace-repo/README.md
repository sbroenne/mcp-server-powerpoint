# PowerPointMcp Agent Plugins

Windows-only Agent Plugins for PowerPointMcp.

This repository is the publish target for plugin artifacts from [`sbroenne/mcp-server-powerpoint`](https://github.com/sbroenne/mcp-server-powerpoint).

> [!WARNING]
> This repository is generated publication output. Do not edit it directly:
> publication overwrites unsynchronized changes. Update the canonical files in
> the [source repository](https://github.com/sbroenne/mcp-server-powerpoint) and
> follow its [plugin publication guide](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates).

## Plugins

- **powerpoint-mcp** — MCP server plugin for conversational PowerPoint automation
- **powerpoint-cli** — CLI plugin for scripting and coding-agent workflows

## Repository Layout

```text
.github/plugin/marketplace.json
plugins/
├── powerpoint-mcp/
│   ├── plugin.json
│   ├── mcp.json
│   └── skills/powerpoint-mcp/SKILL.md
└── powerpoint-cli/
    ├── plugin.json
    └── skills/powerpoint-cli/SKILL.md
```

The canonical marketplace manifest lives at `.github/plugin/marketplace.json`. The `plugins/` directory contains Agent Plugins 1.0 packages generated from source-owned templates by the source repo's `publish-plugins.yml` workflow.

## Install

```powershell
# Register this marketplace
copilot plugin marketplace add sbroenne/mcp-server-powerpoint-plugins

# Install one or both plugins
copilot plugin install powerpoint-mcp@mcp-server-powerpoint-plugins
copilot plugin install powerpoint-cli@mcp-server-powerpoint-plugins
```

Both plugins use the public npm packages through `npx`. Node.js 18 or later is
required.

## Notes

- **Windows only** — PowerPointMcp depends on Microsoft PowerPoint COM automation.
- **powerpoint-mcp** includes portable root `mcp.json` configuration that launches `@sbroenne/mcp-server-powerpoint`.
- **powerpoint-cli** includes an argument-safe npx wrapper for `@sbroenne/pptcli`; separate PATH installation is optional.
- Both root `plugin.json` manifests target `https://agent-plugins.org/schemas/1.0.0/plugin.schema.json`; skills are discovered from the fixed `skills/` directory.

## Source and Support

- Source repo: [sbroenne/mcp-server-powerpoint](https://github.com/sbroenne/mcp-server-powerpoint)
- Issues: [sbroenne/mcp-server-powerpoint/issues](https://github.com/sbroenne/mcp-server-powerpoint/issues)
- Plugin docs: [powerpoint-mcp](https://github.com/sbroenne/mcp-server-powerpoint-plugins/tree/main/plugins/powerpoint-mcp), [powerpoint-cli](https://github.com/sbroenne/mcp-server-powerpoint-plugins/tree/main/plugins/powerpoint-cli)

## License

MIT
