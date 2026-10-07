---
"powerpointmcp": patch
---

**Standalone downloads work with PowerPoint again** (#103): the MCP Server and CLI `.exe` downloads, and the npm packages and Claude Desktop bundle that use them, failed every PowerPoint action with "Built-in COM has been disabled via a feature switch". They are no longer trimmed, so creating and opening presentations works again. The downloads are larger as a result. The NuGet .NET tools were not affected.
