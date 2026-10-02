# PowerPoint CLI Agent Plugin

Windows-only PowerPoint automation for coding agents through `pptcli`.

The plugin installs the `powerpoint-cli` Agent Skill and an argument-safe wrapper that runs
the public `@sbroenne/pptcli` npm package through `npx`.

## Prerequisites

- Windows 10 or later
- Microsoft PowerPoint desktop
- Node.js 18 or later with npm/npx
- Network access for npm installation and updates

## Installation

```text
copilot plugin install sbroenne/mcp-server-powerpoint-plugins/powerpoint-cli
```

## What it supports

- Presentation create, open, test, list, and save-on-close
- Slides, sections, comments, backgrounds, and imports
- Shapes, placeholders, text formatting, and hyperlinks
- Tables, native charts, SmartArt, images, and speaker notes
- Layouts, masters, page setup, accessibility, and animation
- PDF and image export for delivery and visual verification

## Typical workflow

```powershell
$session = npx -y @sbroenne/pptcli@latest session create C:\Decks\demo.pptx | ConvertFrom-Json
npx -y @sbroenne/pptcli@latest slide add-blank -s $session.sessionId
npx -y @sbroenne/pptcli@latest shape add-text-box -s $session.sessionId --slide-index 1 `
  --left 50 --top 50 --width 600 --height 80
npx -y @sbroenne/pptcli@latest session close $session.sessionId --save
```

Run `npx -y @sbroenne/pptcli@latest --help` for the authoritative command
surface.

The wrapper preserves quoted JSON arguments when called from Windows PowerShell. npm selects the
native Windows x64 or ARM64 runtime that matches Node.js.

The plugin contains no PowerPoint documents and does not upload presentation content.
