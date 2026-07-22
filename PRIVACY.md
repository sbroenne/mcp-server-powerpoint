# Privacy Policy

**Last Updated:** July 15, 2026

## Overview

MCP Server for PowerPoint ("PowerPointMcp") is an open-source tool that enables AI assistants and automation scripts to interact with Microsoft PowerPoint. This privacy policy explains how the software handles your data.

## Data Collection Summary

PowerPointMcp processes presentations locally on your machine through Microsoft PowerPoint's desktop COM API.

### What We DO NOT Collect

- ❌ **Presentation contents** - We never collect text, slide content, speaker notes, images, charts, or embedded objects from your PowerPoint files.
- ❌ **File names or paths** - File names and local paths are not collected by the project.
- ❌ **Personal information** - No names, emails, or account information are required.
- ❌ **User accounts** - No registration or sign-in is required.

### Local Processing

PowerPointMcp operates on your local machine:

1. **Local Processing** - PowerPoint operations are performed locally via Microsoft's COM API.
2. **Your Files Stay Local** - Presentation files are read from and written to your local filesystem only.
3. **No Required Cloud Service** - The MCP Server and CLI do not require a hosted PowerPointMcp service.

## Data Flow

When you use PowerPointMcp with an AI assistant:

1. You send a request to the AI assistant.
2. The AI assistant calls PowerPointMcp tools on your local machine.
3. PowerPointMcp performs the requested PowerPoint operations locally.
4. Results are returned to the AI assistant.

**Note:** The AI assistant you use (for example, GitHub Copilot, Claude, or ChatGPT) has its own privacy policy governing how it handles your conversations and data. PowerPointMcp only handles the local PowerPoint automation requested through MCP or CLI commands.

## Third-Party Services

- **Microsoft PowerPoint** - PowerPointMcp requires Microsoft PowerPoint installed on your machine. PowerPoint is subject to Microsoft's privacy policy.
- **AI Assistants** - When used with AI assistants, those services have their own privacy policies.
- **GitHub** - The source code, issues, and releases are hosted on GitHub.

## Open Source

PowerPointMcp is open source software. You can review the complete source code at:
https://github.com/sbroenne/mcp-server-powerpoint

Project documentation is available at:
https://powerpointmcpserver.dev/

## Security

- PowerPointMcp runs with the same permissions as your user account.
- It can only access files and PowerPoint instances that your user account can access.
- No elevated privileges are required or requested.

## Children's Privacy

PowerPointMcp does not knowingly collect any information from anyone, including children under 13 years of age.

## Changes to This Policy

If we make changes to this privacy policy, we will update the "Last Updated" date above and publish the updated policy in our GitHub repository and documentation site.

## Contact

For questions about this privacy policy or the PowerPointMcp project:

- **GitHub Issues:** https://github.com/sbroenne/mcp-server-powerpoint/issues
- **Repository:** https://github.com/sbroenne/mcp-server-powerpoint
- **Documentation:** https://powerpointmcpserver.dev/

---

**Summary:** PowerPointMcp processes your PowerPoint presentations locally on your machine. It does not collect presentation contents, file names, file paths, or personal information.
