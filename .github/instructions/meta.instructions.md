---
applyTo: "AGENTS.md,CLAUDE.md,.github/copilot-instructions.md,.github/instructions/**,.github/github-app.yml,docs/AGENT-CONFIGURATION.md"
excludeAgent: "code-review"
---

# Instruction maintenance

- Keep always-loaded guidance short: repository constraints, required checks,
  non-obvious pitfalls, and source pointers. Avoid tutorials and inventories.
- Give each rule one authoritative home. Put task-specific details in scoped
  `*.instructions.md` files with quoted `applyTo`; link them from the root.
- Agent-specific entry points only import or link shared guidance. Document
  client-specific loading differences instead of maintaining separate rules.
- Include owning generators in scope when generated behavior is involved.
  Implementation instructions exclude `code-review`; review instructions
  exclude `cloud-agent` and keep the essential review checks independently.
- Preserve unique safeguards and update inbound links when consolidating.
  Compare removed content against its destination. Move useful explanations,
  examples, and procedures to linked docs rather than dropping them; document
  corrections to obsolete claims. Keep the migration map in
  `docs/DEVELOPMENT-GUIDE.md` current and agent setup in
  `docs/AGENT-CONFIGURATION.md`.
- Use `.github/github-app.yml` for app settings, not duplicate instructions or
  an unrequested model override. Keep scripts manual unless startup behavior is
  explicitly needed. Do not change account permissions or firewall settings.
