# PowerPoint MCP Server Skill

Agent Skill for AI assistants using the PowerPoint MCP Server via the Model Context Protocol.

## Best For

- **Conversational AI** (Claude Desktop, VS Code Chat, GitHub Copilot Chat)
- Building and editing PowerPoint decks with iterative, visually-verified reasoning
- Multi-slide deck authoring, chart/table/image insertion, speaker notes
- Long-running autonomous tasks that need to inspect and re-inspect visual output

## Installation

### GitHub Copilot / VS Code

Enable skills in VS Code settings:
```json
{
  "chat.useAgentSkills": true
}
```

### Other Platforms

Extract to your AI assistant's skills directory:

| Platform | Location |
|----------|----------|
| **Claude Code** | `.claude/skills/powerpoint-mcp/` |
| **Cursor** | `.cursor/skills/powerpoint-mcp/` |
| **Windsurf** | `.windsurf/skills/powerpoint-mcp/` |
| **Gemini CLI** | `.gemini/skills/powerpoint-mcp/` |
| **Codex** | `.codex/skills/powerpoint-mcp/` |
| **Goose** | `.goose/skills/powerpoint-mcp/` |

Or use npx:
```bash
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

## Contents

```
powerpoint-mcp/
└── SKILL.md           # Compact MCP entry skill
```

Detailed operational guidance is available on the
[documentation site](https://powerpointmcpserver.dev/reference/). The live MCP tool schemas are
the source of truth for available actions and arguments.

## MCP Server Setup

The skill works with the PowerPoint MCP Server (`Sbroenne.PowerPointMcp.McpServer`), a stdio MCP
host that drives a live PowerPoint desktop instance via COM. See the repository README for setup
instructions.

## Related

- [Documentation](https://powerpointmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-powerpoint)
