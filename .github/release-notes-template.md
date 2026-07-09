## PowerPointMcp {{VERSION}}

### What's New
{{CHANGELOG}}

### Installation Options

**VS Code Extension** (Recommended)
- Search "PowerPoint MCP" in VS Code Marketplace and click Install
- Or download `powerpoint-mcp-{{VERSION}}.vsix` below
- Self-contained: no .NET runtime or SDK required
- Auto-registers the MCP server via `mcpServerDefinitionProviders`
- Agent skills (powerpoint-mcp) registered automatically via `chatSkills`

**Claude Desktop (MCPB)**
- Download `powerpoint-mcp-{{VERSION}}.mcpb` and double-click to install

**Standalone Executables** (Primary — no .NET runtime required)
- MCP Server: Download `PowerPointMcp-MCP-Server-{{VERSION}}-windows.zip`, extract `mcp-powerpoint.exe`
- CLI: Download `PowerPointMcp-CLI-{{VERSION}}-windows.zip`, extract `powerpointcli.exe`
- Add the exe(s) to your PATH, then configure your MCP client with command `mcp-powerpoint`

**NuGet (.NET Tool)** (Secondary — requires .NET 10 runtime)
```powershell
dotnet tool install --global Sbroenne.PowerPointMcp.McpServer
dotnet tool install --global Sbroenne.PowerPointMcp.CLI
```

**Agent Skills** (for AI coding assistants)
- VS Code Extension includes the skill automatically (powerpoint-mcp)
- Install via Skills CLI: `npx skills add sbroenne/mcp-server-powerpoint`
- Or download `powerpoint-skills-v{{VERSION}}.zip`

**MCP Registry**
- Discoverable in the official MCP Registry as `io.github.sbroenne/mcp-server-powerpoint`

### Requirements
- Windows OS
- Microsoft PowerPoint 2016+
- No .NET runtime required for VS Code Extension, MCPB, or standalone executables
- .NET 10 Runtime required for NuGet (.NET tool) installation only

### Documentation
- [Website](https://powerpointmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-powerpoint)
- [Changelog](https://github.com/sbroenne/mcp-server-powerpoint/blob/main/CHANGELOG.md)
