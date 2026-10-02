# PowerPointMcp MCP Server

Run the self-contained PowerPointMcp server through npm:

```powershell
npx -y @sbroenne/mcp-server-powerpoint@latest
```

The package requires Node.js 18+, Windows, Microsoft PowerPoint 2016 or later, and an
interactive desktop. It does
not require the .NET SDK or a separately installed .NET runtime.
ARM64 Node.js selects the native ARM64 runtime; x64 Node.js selects x64,
including on ARM64 Windows. Keep optional dependencies enabled. A missing
matching package is an error, not a fallback to another architecture.

The Node.js entry point only launches the packaged .NET server. MCP tools and
PowerPoint automation continue to run in the existing PowerPointMcp implementation.

`@latest` selects the current npm release at startup using normal npm caching.
Network access is needed for downloads and update checks. Restart your MCP
server after safely finishing presentation work to run an updated version.

[Documentation](https://powerpointmcpserver.dev/installation-mcp-server/) |
[Source](https://github.com/sbroenne/mcp-server-powerpoint) |
[Issues](https://github.com/sbroenne/mcp-server-powerpoint/issues)
