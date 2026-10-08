<#
.SYNOPSIS
    Builds the PowerPoint Agent Skills package from canonical skill sources.
#>
[CmdletBinding()]
param(
    [string]$OutputDir = 'artifacts/skills',

    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path $PSScriptRoot -Parent
$SkillsDir = Join-Path $RepoRoot 'skills'
$SkillNames = @('powerpoint-mcp', 'powerpoint-cli', 'powerpoint-deck-design')

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw 'Version is required.'
}

$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDir)) {
    $OutputDir
}
else {
    Join-Path $RepoRoot $OutputDir
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

$staging = Join-Path ([System.IO.Path]::GetTempPath()) "powerpoint-skills-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $staging -Force | Out-Null
try {
    $stagedSkills = Join-Path $staging 'skills'
    New-Item -ItemType Directory -Path $stagedSkills -Force | Out-Null

    foreach ($skillName in $SkillNames) {
        $source = Join-Path $SkillsDir $skillName
        $skillFile = Join-Path $source 'SKILL.md'
        if (-not (Test-Path -LiteralPath $skillFile -PathType Leaf)) {
            throw "Canonical Agent Skill not found: $skillFile"
        }

        $destination = Join-Path $stagedSkills $skillName
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
        Copy-Item -LiteralPath $skillFile -Destination $destination
        [System.IO.File]::WriteAllText(
            (Join-Path $destination 'VERSION'),
            $Version,
            [System.Text.UTF8Encoding]::new($false))
    }

    Copy-Item -LiteralPath (Join-Path $SkillsDir 'README.md') -Destination $staging

    $zipPath = Join-Path $outputPath "powerpoint-skills-v$Version.zip"
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal
}
finally {
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force
    }
}

Copy-Item -LiteralPath (Join-Path $SkillsDir 'CLAUDE.md') -Destination $outputPath -Force
Copy-Item -LiteralPath (Join-Path $SkillsDir '.cursorrules') -Destination $outputPath -Force

$manifest = [ordered]@{
    name = 'powerpoint-skills'
    version = $Version
    description = 'PowerPoint MCP Server Agent Skills for AI coding assistants'
    platforms = @('github-copilot', 'claude-code', 'cursor', 'windsurf', 'gemini-cli', 'goose', 'codex')
    skills = @(
        [ordered]@{ name = 'powerpoint-mcp'; path = 'skills/powerpoint-mcp'; target = 'MCP Server' }
        [ordered]@{ name = 'powerpoint-cli'; path = 'skills/powerpoint-cli'; target = 'CLI Tool' }
        [ordered]@{ name = 'powerpoint-deck-design'; path = 'skills/powerpoint-deck-design'; target = 'Deck Design' }
    )
    repository = 'https://github.com/sbroenne/mcp-server-powerpoint'
    documentation = 'https://powerpointmcpserver.dev/'
}
$manifestJson = ($manifest | ConvertTo-Json -Depth 10) -replace "`r?`n", "`n"
[System.IO.File]::WriteAllText(
    (Join-Path $outputPath 'manifest.json'),
    "$manifestJson`n",
    [System.Text.UTF8Encoding]::new($false))

Write-Output "Built PowerPoint Agent Skills package v$Version in $outputPath."
