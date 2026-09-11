# Copilot configuration

GitHub Copilot is the primary coding agent locally and on GitHub. Both use
[the repository instructions](../.github/copilot-instructions.md).
[AGENTS.md](../AGENTS.md) provides the same entry point to tools that discover
that filename, including Codex. [CLAUDE.md](../CLAUDE.md) imports the shared root
and critical rules for Claude Code. No custom agent or fixed model is required.

## Other agents

| Client | Entry point | Task-specific rules |
| --- | --- | --- |
| Copilot locally and on GitHub | `.github/copilot-instructions.md` | Scoped instruction files plus root links |
| Codex and other `AGENTS.md` readers | `AGENTS.md` | Follow the shared root's task links |
| Claude Code | `CLAUDE.md` imports | Follow the shared root's task links |
| Clients with another instruction format | Point their project guidance at `AGENTS.md` | Read the linked files explicitly |

GitHub's `applyTo` and `excludeAgent` metadata is not portable to every client.
The shared entry points therefore instruct agents to read relevant files, rather
than assuming automatic loading. The app commands and cloud setup workflow below
are Copilot-specific; other agents use their own environment setup. These files
provide guidance, not tool permissions or a guarantee that a client obeys it.

## Where guidance belongs

- `.github/copilot-instructions.md`: short shared entry point and task links.
- `.github/instructions/critical-rules.instructions.md`: always-needed safety rules.
- Other `.github/instructions/*.instructions.md`: task-specific guidance selected
  by file path. Root links also make it available when a client does not load
  path-specific files automatically.
- `code-review.instructions.md`: independent review checks, not implementation
  procedures. Implementation files exclude code review to avoid duplicate advice.
- `.github/github-app.yml`: local app commands, without another copy of the rules.
- `.github/workflows/copilot-setup-steps.yml`: GitHub coding-agent environment setup.

This follows the instruction structure in `sbroenne/mcp-server-excel`, adapted
for PowerPoint's PIA embedding, process ownership, and existing test filters.
Keep detailed architecture in its scoped guide rather than growing the root
prompt with operation counts or repeated examples.
The [development guide](DEVELOPMENT-GUIDE.md) preserves the supporting
explanations, procedures, and a section-by-section map of the former guidance.

## Local app

Review and accept `.github/github-app.yml` when prompted by the Copilot app.
Repository configuration changes are not applied until accepted. The commands
are manual: opening a session does not start builds, install dependencies, launch
PowerPoint, or terminate processes.

The build and PowerPoint-free test commands disable build-triggered process
cleanup so they do not interrupt other live sessions. Test hang detection is a
backstop; agent-driven test runs must also have an overall execution timeout.
For Core/COM work, run the focused real-PowerPoint tests described in
[the testing strategy](../.github/instructions/testing-strategy.instructions.md).
The normal pre-commit hook remains authoritative; these commands do not replace it.

## GitHub coding agent

The setup workflow takes effect for agent sessions after it reaches the default
branch. It can also be run manually or checked on a PR changing that workflow.
It prepares Windows, the SDK from `global.json`, Node.js, Python, and dependencies
for the solution, extension, and documentation. It does not install PowerPoint.

Windows coding-agent environments do not support Copilot's integrated firewall.
Before using this workflow for cloud sessions, the repository administrator must
review runner availability and network controls. GitHub recommends a self-hosted
or larger Windows runner with appropriate network controls. If required, replace
`runs-on` with the approved Windows runner label and configure the repository
settings separately. This change does not provision runners, change permissions,
or disable the firewall.

On a Windows runner without PowerPoint, builds, static audits, packaging tests,
and MCP tests filtered with `RequiresPowerPoint!=true` can run. Core and
PowerPoint-dependent lifecycle tests require local desktop PowerPoint; report
them as not run when unavailable. Do not replace them with mocked COM tests.
Use the existing [CI Gate](../.github/workflows/ci.yml) as the build/check reference.

Setup failures are visible in the workflow logs; the agent may start with only
part of the environment ready. Check the actual tool availability rather than
assuming setup succeeded.

References:
[app configuration](https://docs.github.com/en/copilot/reference/github-copilot-app-reference/repository-configuration)
and [cloud environment setup](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/customize-the-agent-environment).
