# PowerPoint MCP Server - Agent Skills

| Skill | Component | Distribution | Best For |
|-------|-----------|---------------|----------|
| **[powerpoint-mcp](powerpoint-mcp/SKILL.md)** | MCP Server (`mcp-powerpoint.exe`) | GitHub Release, NuGet .NET tool, VS Code extension, MCPB, direct skill extraction | Conversational AI — rich MCP tool schemas |

**Shared guidance:** `skills/shared/*.md` — single source of truth, copied into
`skills/powerpoint-mcp/references/`. A future `powerpoint-cli` skill (once the CLI is built out
with real Generators-based commands, matching `mcp-server-excel`'s `excel-cli`) would draw from
the same shared guidance.

## Installation

**Direct skill extraction (for agents without plugin support):**
```bash
npx skills add sbroenne/mcp-server-powerpoint --skill powerpoint-mcp
```

**Via VS Code Extension** (planned — see `.squad/decisions.md` release-deliverables entry):
Installs this skill automatically to `~/.copilot/skills/powerpoint-mcp/`.

## Building

There is no automated skill-build step yet (no Generators/build script for this project — see
`.github/copilot-instructions.md`). `skills/powerpoint-mcp/references/*.md` is currently a manual
copy of `skills/shared/*.md`; re-copy after editing shared content until a build step is added.

## Structure

```
skills/
├── shared/          # Shared authoring guidance (source of truth)
├── powerpoint-mcp/  # MCP Server skill (SKILL.md + references/)
├── CLAUDE.md        # Claude Code project instructions
└── .cursorrules     # Cursor-specific rules
```
