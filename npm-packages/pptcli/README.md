# PowerPointMcp CLI

Run the self-contained PowerPoint automation CLI through npm:

```powershell
npx -y @sbroenne/pptcli@latest --help
npx -y @sbroenne/pptcli@latest -q session open "C:\Data\Deck.pptx"
```

For repeated use, install the `pptcli` command on your PATH:

```powershell
npm install --global @sbroenne/pptcli@latest
pptcli --help
```

Requires Node.js 18 or later, Windows x64 or ARM64, and
Microsoft PowerPoint 2016 or later on an interactive desktop. No separate .NET runtime is needed. Keep
optional dependencies enabled so npm installs the matching Windows runtime.
ARM64 Node.js selects the native ARM64 runtime; x64 Node.js selects x64,
including on ARM64 Windows. A missing matching package is an error, not a
fallback to another architecture.

The launcher forwards arguments, standard input/output, and exit codes to the
existing CLI. PowerPoint operations and session management are unchanged.

`@latest` selects the current npm release using normal npm caching. It does not
replace an already running background service. Finish and explicitly save/close
presentation sessions before stopping that service to use a new version.
Network access is needed for package downloads and update checks.

[Documentation](https://powerpointmcpserver.dev/installation-cli/) |
[Source](https://github.com/sbroenne/mcp-server-powerpoint) |
[Issues](https://github.com/sbroenne/mcp-server-powerpoint/issues)
