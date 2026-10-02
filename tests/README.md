# PowerPointMcp Tests

PowerPoint behavior uses real desktop PowerPoint through COM. Protocol,
generation, packaging, and CLI mapping can run without PowerPoint.

## PowerPoint-free checks

After a Release build:

```powershell
.\scripts\Invoke-PowerPointFreeTests.ps1
```

This runs the CLI, MCP protocol, and SkillGeneration projects with a hard
process deadline and verifies that every selected group produced and passed at
least one test. MCP tests marked `RequiresPowerPoint=true` are excluded.

For a focused local contract check:

```powershell
.\scripts\Invoke-PowerPointFreeTests.ps1 -Local -Contracts
```

## Real PowerPoint checks

Core and ComInterop behavior must use real PowerPoint and stay serialized:

```powershell
dotnet test tests\PowerPointMcp.Core.Tests --filter "Feature=Shape"
dotnet test tests\PowerPointMcp.McpServer.Tests
```

Use an explicit command timeout. The complete MCP project includes session
round trips marked `RequiresPowerPoint=true`; the PowerPoint-free runner does
not claim those passed.

Do not add mocked COM command tests or enable parallel execution. Protocol tests
should assert behavior owned by this repository: tool/action completeness,
declared parameter metadata, injected-parameter exclusion, structured results,
output fields, validation, and actual dispatch. Avoid locking tests to SDK-owned
primitive JSON Schema details.

## Changed-file planning

`scripts/Get-ValidationPlan.ps1` maps changed paths to affected builds,
PowerPoint tests, packages, skills, plugins, MCPB, and VSIX outputs. Pre-commit
and CI use that shared classification instead of maintaining separate path
rules.
