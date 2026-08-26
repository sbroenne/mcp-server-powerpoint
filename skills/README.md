# PowerPoint MCP Server - Agent Skills

| Skill | Component | Distribution | Best For |
|-------|-----------|---------------|----------|
| **[powerpoint-mcp](powerpoint-mcp/SKILL.md)** | MCP Server (`mcp-powerpoint.exe`) | Agent Plugin, GitHub Release, VS Code extension, MCPB, direct skill extraction | Conversational AI — rich MCP tool schemas |
| **[powerpoint-cli](powerpoint-cli/SKILL.md)** | CLI (`powerpointcli.exe`) | Agent Plugin, GitHub Release, direct skill extraction | Coding agents and scripts — compact command surface |

**Shared guidance:** `skills/shared/*.md` is the single source of truth copied into both skills
during deterministic packaging. The CLI package also receives a command reference generated from
live recursive help.

## Installation

**Direct skill extraction (for agents without plugin support):**
```bash
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-cli
```

**Via VS Code Extension:**
Installs the MCP skill automatically to `~/.copilot/skills/powerpoint-mcp/`.

## Building

`scripts/Build-AgentSkills.ps1` builds both skill packages, synchronizes shared guidance, generates
the CLI command reference from live help, stamps package-only versions, and creates the release
archive.

## Structure

```
skills/
├── shared/          # Shared authoring guidance (source of truth)
├── powerpoint-mcp/  # MCP Server skill (SKILL.md + references/)
├── powerpoint-cli/  # CLI skill (SKILL.md + generated command reference)
├── CLAUDE.md        # Claude Code project instructions
└── .cursorrules     # Cursor-specific rules
```
