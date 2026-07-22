# Security Policy

## Supported Versions

PowerPointMcp ships frequent releases. We only support the **latest published version** with security fixes — there are no parallel maintenance branches for older minor/patch releases:

| Version              | Supported          |
| -------------------- | ------------------ |
| Latest release       | :white_check_mark: |
| Any older release    | :x: Please upgrade |

Check the [Releases page](https://github.com/sbroenne/mcp-server-powerpoint/releases) for the current latest version, and keep your installation up to date via the CLI, MCPB bundle, VS Code extension, or NuGet.

## Security Features

PowerPointMcp includes several security measures:

### Input Validation

- **Path Traversal Protection**: File paths are resolved with `Path.GetFullPath()` before use.
- **Extension Validation**: PowerPoint presentation files such as `.pptx` and `.pptm` are the intended inputs.
- **Path Length Validation**: Windows path limits are respected.
- **Session Validation**: MCP and CLI operations validate session identifiers before dispatching commands.

### Code Analysis

- **Enhanced Security Rules**: Security-focused .NET analyzers are enforced in the repo configuration.
- **Treat Warnings as Errors**: Code quality issues must be resolved before release.
- **CodeQL Scanning**: Automated security scanning runs through GitHub Actions.

### COM Security

- **Controlled PowerPoint Automation**: PowerPoint is automated locally through the official desktop COM API.
- **Resource Cleanup**: COM objects and presentation sessions are disposed through the shared session lifecycle.
- **No Remote Connections**: Only local PowerPoint automation is supported.

### PowerPointMcp Service Security

The PowerPointMcp service manages PowerPoint COM automation sessions:

**MCP Server**: The service runs fully **in-process** — no inter-process communication. There is no attack surface beyond the MCP Server process itself.

**CLI**: The CLI daemon uses a **Windows named pipe** for communication between CLI commands and the daemon process:

| Protection | Status | Description |
| ---------- | ------ | ----------- |
| **User Isolation** | ✅ Enforced | Pipe access is scoped to the current Windows user. |
| **Windows ACLs** | ✅ Enforced | Named pipe security restricts access to the current user. |
| **Local Only** | ✅ Enforced | Named pipes are local IPC only; no network access is possible. |
| **Process Restriction** | ❌ Not Enforced | Any process running as the same user can connect to the CLI daemon. |

**What This Means:**

1. **Same-user access**: Any application running under your Windows user account can connect to the CLI daemon and execute PowerPoint operations. This is by design, similar to local developer daemons.
2. **No cross-user access**: Other Windows users cannot connect to your CLI daemon.
3. **No network access**: The named pipe is strictly local. Remote processes cannot connect.

### Dependency Management

- **Dependabot**: Automated dependency updates and security patches.
- **Dependency Review**: Pull request scanning for vulnerable dependencies.
- **Central Package Management**: Consistent dependency versioning across all projects.

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security issue, please follow these steps:

### 1. **DO NOT** Create a Public Issue

Please do not create a public GitHub issue for security vulnerabilities. This could put all users at risk.

### 2. Report Privately

Report security vulnerabilities using one of these methods:

**Preferred Method: GitHub Security Advisories**

1. Go to <https://github.com/sbroenne/mcp-server-powerpoint/security/advisories>
2. Click "Report a vulnerability"
3. Fill out the advisory form with detailed information

**Alternative: GitHub Direct Message**

Contact the maintainer via GitHub: [@sbroenne](https://github.com/sbroenne)

Subject: `[SECURITY] PowerPointMcp Vulnerability Report`

### 3. Information to Include

Please provide as much information as possible:

- **Description**: Clear description of the vulnerability
- **Impact**: What could an attacker do with this vulnerability?
- **Affected Versions**: Which versions are affected?
- **Proof of Concept**: Steps to reproduce, if possible
- **Suggested Fix**: If you have a fix or mitigation

### 4. What to Expect

- **Acknowledgment**: Within 48 hours
- **Initial Assessment**: Within 5 business days
- **Status Updates**: Regular updates on progress
- **Fix Timeline**:
  - Critical: 7 days
  - High: 30 days
  - Medium: 90 days
  - Low: Best effort

### 5. Coordinated Disclosure

We follow responsible disclosure practices:

1. **Private Fix**: We develop a fix privately.
2. **Security Advisory**: We create a GitHub Security Advisory.
3. **CVE Assignment**: We request a CVE if applicable.
4. **Public Release**: We release the patch with security notes.
5. **Credit**: We credit reporters in release notes if desired.

## Security Best Practices for Users

### MCP Server Security

- **Validate AI Requests**: Review PowerPoint operations requested by AI assistants.
- **File Path Restrictions**: Only allow MCP Server access to directories you trust.
- **Audit Logs**: Monitor MCP Server operations in logs.
- **Sensitive Decks**: Avoid exposing presentations with sensitive data to untrusted AI assistants.

### CLI Security

- **Script Validation**: Review automation scripts before execution.
- **File Permissions**: Ensure presentation files have appropriate permissions.
- **Isolated Environment**: Run in a sandboxed environment when processing untrusted files.
- **PowerPoint Security Settings**: Maintain appropriate Office macro and add-in security settings.

### Development Security

- **Code Review**: All changes require review before merge.
- **Branch Protection**: Main branch should be protected with required checks.
- **Signed Commits**: Consider using signed commits.
- **Least Privilege**: Run with minimal required permissions.

## Known Security Considerations

### PowerPoint COM Automation

- **Local Only**: PowerPointMcp only supports local PowerPoint automation.
- **Windows Only**: Requires Windows with Microsoft PowerPoint installed.
- **PowerPoint Process**: Creates PowerPoint COM objects under the current user account.
- **Macro Security**: Macro-enabled presentations remain subject to PowerPoint Trust Center settings.

### File System Access

- **Full Path Resolution**: Paths are resolved to absolute paths before use.
- **Current User Context**: Operations run with current user permissions.
- **Local Files**: Use trusted local files and directories for automation.

### AI Integration (MCP Server)

- **Trusted AI Assistants**: Only use with trusted AI platforms.
- **Request Validation**: Review operations before PowerPoint executes them.
- **Sensitive Data**: Avoid exposing presentations with sensitive data to AI assistants.
- **Audit Trail**: MCP Server logs operations for diagnostics.

## Security Updates

Security updates are published through:

- **GitHub Security Advisories**: <https://github.com/sbroenne/mcp-server-powerpoint/security/advisories>
- **Release Notes**: <https://github.com/sbroenne/mcp-server-powerpoint/releases>
- **Project Site**: <https://powerpointmcpserver.dev/>

Subscribe to repository notifications to receive security alerts.

## Vulnerability Disclosure Policy

### Our Commitment

- We acknowledge receipt of vulnerability reports within 48 hours.
- We keep reporters informed of progress.
- We credit researchers in security advisories if desired.
- We do not take legal action against researchers following responsible disclosure.

### Researcher Guidelines

- **Responsible Disclosure**: Give us time to fix before public disclosure.
- **No Harm**: Do not access, modify, or delete other users' data.
- **Good Faith**: Act in good faith to help improve security.
- **Legal**: Follow all applicable laws.

## Security Contacts

- **GitHub Security**: <https://github.com/sbroenne/mcp-server-powerpoint/security>
- **Maintainer**: @sbroenne

## Additional Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Microsoft Security Response Center](https://msrc.microsoft.com/)
- [CVE Database](https://cve.mitre.org/)
- [National Vulnerability Database](https://nvd.nist.gov/)

---

**Last Updated**: 2026-07-15

Thank you for helping keep PowerPointMcp and its users safe!
