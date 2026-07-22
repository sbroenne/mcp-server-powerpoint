---
name: MCP Server Issue
about: Report issues with the MCP Server for AI assistants
title: '[MCP] '
labels: 'mcp-server'
assignees: ''

---

## Issue Description
A clear and concise description of the MCP Server issue.

## AI Assistant
Which AI assistant are you using with the MCP Server?
- [ ] **GitHub Copilot** (VS Code, Visual Studio, etc.)
- [ ] **Claude Desktop** (Anthropic)
- [ ] **ChatGPT** (OpenAI)
- [ ] **Other**: [please specify]

## MCP Tool & Action
Which MCP tool and action are experiencing issues?
- **Tool**: [e.g., presentation, slide, shape, textframe, chart, export]
- **Action**: [e.g., create, open, add-slide, add-shape, export-slide]
- **File Path**: [e.g., `C:\Decks\presentation.pptx`]
- **Additional Parameters**: [describe any other parameters used]

## Expected Behavior
What did you expect the MCP Server to do?

## Actual Behavior
What did the MCP Server actually do?

## Error Response
If you received an error, paste the full JSON response:
```json
{
  "error": "paste error here"
}
```

## MCP Server Configuration
How is the MCP Server configured?

**Configuration file location**: [e.g., `.config/Code/User/globalStorage/github.copilot-chat/config.json`]

**MCP Configuration**:
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "powerpoint-mcp-server"
    }
  }
}
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **PowerPoint Version**: [e.g. PowerPoint 365, PowerPoint 2021]
- **PowerPointMcp Version**: [e.g. v0.1.0]
- **.NET Version**: [Run `dotnet --version`]
- **Installation Method**:
  - [ ] MCPB bundle
  - [ ] Global .NET tool
  - [ ] Binary download
  - [ ] Source build
  - [ ] Other: [please specify]

## MCP Server Logs
If possible, provide relevant logs from the MCP Server:
```
[Paste logs here]
```

## Steps to Reproduce
1. Configure AI assistant with MCP Server
2. Ask AI assistant: "..."
3. MCP Server receives request for tool: [tool_name], action: [action_name]
4. See error

## Conversation Context (Optional)
If helpful, provide the conversation you had with the AI assistant that led to this issue:
```
User: "Can you create a presentation with three slides?"
AI: [response]
[MCP Server error occurs]
```

## Presentation Details
- **File Format**: [.pptx or .pptm]
- **File Size**: [approximate size]
- **Contains**:
  - [ ] Multiple slides
  - [ ] Shapes/text boxes
  - [ ] Tables
  - [ ] Charts or SmartArt
  - [ ] Images or media
  - [ ] Speaker notes
  - [ ] Custom layouts or masters

## Additional Context
Add any other context about the problem here, including:
- Screenshots of AI assistant interaction
- Sample PowerPoint files (with sensitive data removed)
- Other relevant information
