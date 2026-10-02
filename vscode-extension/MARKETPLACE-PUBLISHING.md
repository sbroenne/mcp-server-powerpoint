# VS Code Marketplace Publishing

The release workflow publishes both verified packages:

- `powerpoint-mcp-<version>.vsix` (`win32-x64`)
- `powerpoint-mcp-<version>-win32-arm64.vsix` (`win32-arm64`)

Both artifacts are built by `scripts/Build-VscodeExtension.ps1`. Do not rename a
generic VSIX or reuse the same native executable for both targets.

Publishing runs on Windows with the locked `@vscode/vsce` dependency:

```powershell
npm ci --ignore-scripts
npm exec --no -- vsce publish --packagePath <package.vsix> --skip-duplicate
```

`VSCE_PAT` must contain the Marketplace token. Publication failures stop the
release; they are not warning-only. `--skip-duplicate` makes an exact release
retry safe without hiding other failures.
