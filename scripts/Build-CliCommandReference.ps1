#Requires -Version 7.0
<#
.SYNOPSIS
    Regenerates skills/powerpoint-cli/references/cli-commands.md from the generated SKILL.md.

.DESCRIPTION
    skills/powerpoint-cli/SKILL.md is emitted by the CLI source generator during a Release
    build, so it is always an exact reflection of the real command surface. The companion
    references/cli-commands.md was previously hand-maintained and had drifted 29 operations
    behind - including the entire `smartart` domain, which an agent reading it would have
    concluded did not exist.

    This script derives the domain reference from SKILL.md so the two can no longer disagree.
    It is the PowerPoint equivalent of the sibling mcp-server-excel repo's
    Build-AgentSkills.ps1 -PopulateReferences.

    Run it after any Release build that regenerates SKILL.md. scripts/check-doc-counts.ps1
    fails if the checked-in file does not match what this script would produce.

.PARAMETER Check
    Do not write anything; exit 1 if the checked-in file is out of date.
#>
[CmdletBinding()]
param(
    [switch]$Check
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$skillPath = Join-Path $repoRoot 'skills\powerpoint-cli\SKILL.md'
$outputPath = Join-Path $repoRoot 'skills\powerpoint-cli\references\cli-commands.md'

if (-not (Test-Path $skillPath)) {
    throw "SKILL.md not found at '$skillPath'. Run a Release build first so the generator emits it."
}

# Hand-written preamble: the session/service commands are wired up by hand in the CLI rather
# than emitted by the domain generator, so they are not present in SKILL.md's domain sections.
$header = @'
# CLI Command Reference

> Generated from `skills/powerpoint-cli/SKILL.md` by `scripts/Build-CliCommandReference.ps1`.
> Do not edit by hand — run the script instead. Use these exact parameter names.

## Global Commands

### session

Open, create, save, close, or list presentation sessions held by the daemon.

**Commands:** `open <FILE_PATH>`, `create <FILE_PATH>`, `close <SESSION_ID>`, `save <SESSION_ID>`, `list`

| Command | Description |
|---------|-------------|
| `session open <FILE_PATH>` | Open an existing presentation and return a session id |
| `session create <FILE_PATH>` | Create a new presentation and return a session id |
| `session close <SESSION_ID>` | Close a session, optionally saving first |
| `session save <SESSION_ID>` | Save the presentation open in a session |
| `session list` | List every session currently open in the daemon |

### service

Start, stop, or check the status of the `pptcli` background daemon.

**Commands:** `start`, `stop`, `status`

| Command | Description |
|---------|-------------|
| `service start` | Start the daemon if it isn't already running |
| `service stop` | Stop the running daemon |
| `service status` | Report whether the daemon is running |

## Domain Commands

Every domain command below follows the shape `pptcli <domain> <ACTION> [OPTIONS]`, targeting an
already-open session via `-s, --session <SESSION>` (obtained from `session open`/`session
create`). All slide/shape/row/column indices are 1-based, matching PowerPoint's own COM object
model.
'@

# Parse SKILL.md's generated domain sections: "### `domain` — description", an "Actions:" line,
# and a flag table. Everything before the first such heading is skill guidance, not reference.
$lines = Get-Content $skillPath
$sections = [System.Collections.Generic.List[object]]::new()
$current = $null

foreach ($line in $lines) {
    if ($line -match '^###\s+`([a-z]+)`\s*[-\u2014]\s*(.+)$') {
        if ($null -ne $current) { $sections.Add($current) }
        $current = [pscustomobject]@{
            Domain      = $Matches[1]
            Description = $Matches[2].Trim()
            Body        = [System.Collections.Generic.List[string]]::new()
        }
        continue
    }

    if ($null -eq $current) { continue }

    # A non-domain heading ends the generated reference block.
    if ($line -match '^#{1,3}\s' ) {
        $sections.Add($current)
        $current = $null
        continue
    }

    $current.Body.Add($line)
}
if ($null -ne $current) { $sections.Add($current) }

if ($sections.Count -eq 0) {
    throw "No generated domain sections found in '$skillPath'. Has the generator output format changed?"
}

$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine($header)

foreach ($section in $sections) {
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("### $($section.Domain)")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine($section.Description)

    $body = $section.Body -join "`n"
    $body = $body -replace '(?m)^Actions:', '**Actions:**'
    $body = $body -replace '(?m)^\| Flag \| Description \|$', '| Parameter | Description |'
    [void]$sb.AppendLine($body.TrimEnd())
}

$rendered = ($sb.ToString().TrimEnd() -replace "`r`n", "`n") + "`n"

if ($Check) {
    if (-not (Test-Path $outputPath)) {
        Write-Host "MISSING: $outputPath" -ForegroundColor Red
        exit 1
    }
    $existing = ((Get-Content $outputPath -Raw) -replace "`r`n", "`n").TrimEnd() + "`n"
    if ($existing -ne $rendered) {
        Write-Host "OUT OF DATE: skills/powerpoint-cli/references/cli-commands.md" -ForegroundColor Red
        Write-Host "  Run: pwsh -File scripts/Build-CliCommandReference.ps1" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "cli-commands.md is up to date ($($sections.Count) domains)." -ForegroundColor Green
    exit 0
}

$outputDir = Split-Path -Parent $outputPath
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
Set-Content -Path $outputPath -Value $rendered -NoNewline -Encoding utf8
Write-Host "Wrote skills/powerpoint-cli/references/cli-commands.md ($($sections.Count) domains)." -ForegroundColor Green
