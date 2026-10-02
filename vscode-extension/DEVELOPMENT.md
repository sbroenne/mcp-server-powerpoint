# VS Code Extension Development

The extension bundles one self-contained native MCP server and the canonical
PowerPoint MCP skill. Separate VSIX files target Windows x64 and Windows ARM64;
each contains a matching executable.

## Local checks

```powershell
Set-Location vscode-extension
npm ci --ignore-scripts
npm run compile
npm run lint
npm run typecheck:tests
npm test
```

The Vitest suite checks registration, packaged version propagation, launch-time
prerequisites, cancellation, and safe setup diagnostics. It does not launch
PowerPoint.

## Release-shaped packaging

From the repository root:

```powershell
.\scripts\Build-VscodeExtension.ps1 -Version 1.2.3 -OutputDir artifacts\vscode-check
```

Packaging uses an isolated temporary directory rather than writing native
runtimes, generated skills, compiled output, or changelog copies into this
source folder. It:

1. Publishes x64 and ARM64 native MCP servers.
2. Copies and stamps the canonical `powerpoint-mcp` skill.
3. Installs locked npm dependencies and runs compile, lint, type checks, and tests.
4. Packages `win32-x64` and `win32-arm64` VSIX files with locked `vsce`.
5. Opens each finished archive and verifies its target, version, required files,
   skill version, and the actual PE machine type of the bundled server.

Do not package with a server copied manually into `vscode-extension/bin`; the
shared script is the authoritative path.

## Runtime behavior

Discovery returns the bundled executable and extension version without probing
the machine. Immediately before launch, `resolveMcpServerDefinition` verifies
Windows, executable readability, and PowerPoint COM registration with a
cancellable, noninteractive Windows PowerShell check.

The extension is declared as `extensionKind: ["ui"]` because PowerPoint must run
on the user's local interactive Windows desktop.
