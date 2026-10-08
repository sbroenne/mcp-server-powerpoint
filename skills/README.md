# PowerPoint MCP Server - Agent Skills

| Skill | Component | Distribution | Best For |
|-------|-----------|---------------|----------|
| **[powerpoint-mcp](powerpoint-mcp/SKILL.md)** | MCP Server (`mcp-powerpoint.exe`) | Agent Plugin, GitHub Release, VS Code extension, MCPB, direct skill extraction | Conversational AI — rich MCP tool schemas |
| **[powerpoint-cli](powerpoint-cli/SKILL.md)** | CLI (`powerpointcli.exe`) | Agent Plugin, GitHub Release, direct skill extraction | Coding agents and scripts — compact command surface |
| **[powerpoint-deck-design](powerpoint-deck-design/SKILL.md)** | PowerPoint design | Agent Plugin, GitHub Release, direct skill extraction | Optional guidance for visual design and review |

Detailed task guidance lives on the
[PowerPoint MCP Server documentation site](https://powerpointmcpserver.dev/reference/). Skills
stay focused on tool discovery and the small set of safe-use rules; the CLI discovers commands
from live `--help` output instead of bundling a generated catalog.

## Installation

**Direct skill extraction (for agents without plugin support):**
```bash
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-cli
```

**Via VS Code Extension:** Installs the MCP skill automatically to
`~/.copilot/skills/powerpoint-mcp/`.

## Building

`scripts/Build-AgentSkills.ps1` packages the three skills with package-only version files and
creates the release archive. Plugin builds also copy the relevant entry skill and the optional
deck-design skill.

## Structure

```
skills/
├── powerpoint-mcp/  # Compact MCP entry skill
├── powerpoint-cli/  # Compact CLI entry skill
├── powerpoint-deck-design/ # Optional visual-design guidance
├── CLAUDE.md        # Claude Code project instructions
└── .cursorrules     # Cursor-specific rules
```
