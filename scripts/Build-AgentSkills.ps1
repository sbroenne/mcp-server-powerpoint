<#
.SYNOPSIS
    Builds the PowerPoint MCP Agent Skills package from canonical sources.
#>
[CmdletBinding()]
param(
    [string]$OutputDir = 'artifacts/skills',

    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [switch]$PopulateReferences,

    [string]$CliPath
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path $PSScriptRoot -Parent
$SkillsDir = Join-Path $RepoRoot 'skills'
$SharedDir = Join-Path $SkillsDir 'shared'

function Copy-SharedReferences([string]$SkillPath, [switch]$CliSyntax) {
    $references = Join-Path $SkillPath 'references'
    New-Item -ItemType Directory -Path $references -Force | Out-Null

    foreach ($source in Get-ChildItem -LiteralPath $SharedDir -File -Filter '*.md') {
        $destination = Join-Path $references $source.Name
        if ($CliSyntax) {
            $notice = '> **CLI syntax:** Shared guides may use MCP calls as shorthand. Use `cli-commands.md` or live `--help` for exact commands and kebab-case options.'
            "$notice`r`n`r`n$(Get-Content -LiteralPath $source.FullName -Raw)" |
                Set-Content -LiteralPath $destination -Encoding utf8 -NoNewline
        }
        else {
            Copy-Item -LiteralPath $source.FullName -Destination $destination -Force
        }
    }
}

function Get-HelpSection([string[]]$Lines, [string]$Header) {
    $start = [Array]::IndexOf($Lines, $Header)
    if ($start -lt 0) {
        return @()
    }

    $section = [System.Collections.Generic.List[string]]::new()
    for ($index = $start + 1; $index -lt $Lines.Count; $index++) {
        if ($Lines[$index] -match '^[A-Z][A-Z ]+:$') {
            break
        }

        $section.Add($Lines[$index])
    }

    @($section)
}

function Get-HelpCommandNames([string[]]$Lines) {
    @(
        Get-HelpSection $Lines 'COMMANDS:' |
            ForEach-Object {
                if ($_ -match '^\s{4}(?<name>[a-z][a-z0-9-]*)(?:\s+<[^>]+>)?\s{2,}') {
                    $Matches.name
                }
            } |
            Sort-Object -Unique
    )
}

function Invoke-CliHelp([string[]]$Arguments) {
    $previousNoColor = $env:NO_COLOR
    $env:NO_COLOR = '1'
    try {
        $output = @(& $script:ResolvedCliPath @Arguments 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "CLI help failed: $script:ResolvedCliPath $($Arguments -join ' ')"
        }

        @($output | ForEach-Object { "$_".TrimEnd() })
    }
    finally {
        $env:NO_COLOR = $previousNoColor
    }
}

function Generate-CliReference([string]$SkillPath) {
    $mainHelp = Invoke-CliHelp @('--help')
    $content = [System.Collections.Generic.List[string]]::new()
    $content.Add('# CLI Command Reference')
    $content.Add('')
    $content.Add('> Auto-generated recursively from the built `pptcli` runtime. Do not edit by hand.')
    $content.Add('')

    foreach ($command in Get-HelpCommandNames $mainHelp) {
        $commandHelp = Invoke-CliHelp @($command, '--help')
        $content.Add("## ``pptcli $command``")
        $content.Add('')
        $content.Add('```text')
        $content.AddRange([string[]]$commandHelp)
        $content.Add('```')
        $content.Add('')

        foreach ($subcommand in Get-HelpCommandNames $commandHelp) {
            $subcommandHelp = Invoke-CliHelp @($command, $subcommand, '--help')
            $content.Add("### ``pptcli $command $subcommand``")
            $content.Add('')
            $content.Add('```text')
            $content.AddRange([string[]]$subcommandHelp)
            $content.Add('```')
            $content.Add('')
        }
    }

    $references = Join-Path $SkillPath 'references'
    New-Item -ItemType Directory -Path $references -Force | Out-Null
    $content -join "`n" |
        Set-Content -LiteralPath (Join-Path $references 'cli-commands.md') -Encoding utf8 -NoNewline
}

if ([string]::IsNullOrWhiteSpace($CliPath)) {
    $CliPath = Join-Path $RepoRoot 'src\PowerPointMcp.CLI\bin\Release\net10.0-windows\powerpointcli.exe'
}
$script:ResolvedCliPath = $CliPath

if ($PopulateReferences) {
    Copy-SharedReferences (Join-Path $SkillsDir 'powerpoint-mcp')
    Copy-SharedReferences (Join-Path $SkillsDir 'powerpoint-cli') -CliSyntax
    if (-not (Test-Path -LiteralPath $script:ResolvedCliPath -PathType Leaf)) {
        throw "pptcli was not found at '$script:ResolvedCliPath'. Build the Release CLI first."
    }
    Generate-CliReference (Join-Path $SkillsDir 'powerpoint-cli')
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw 'Version is required unless -PopulateReferences is used.'
}
if (-not (Test-Path -LiteralPath $script:ResolvedCliPath -PathType Leaf)) {
    throw "pptcli was not found at '$script:ResolvedCliPath'. Build the Release CLI first."
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

    foreach ($skillName in @('powerpoint-mcp', 'powerpoint-cli')) {
        $destination = Join-Path $stagedSkills $skillName
        Copy-Item -LiteralPath (Join-Path $SkillsDir $skillName) -Destination $destination -Recurse
        Copy-SharedReferences $destination -CliSyntax:($skillName -eq 'powerpoint-cli')
        [System.IO.File]::WriteAllText(
            (Join-Path $destination 'VERSION'),
            $Version,
            [System.Text.UTF8Encoding]::new($false))
    }

    Generate-CliReference (Join-Path $stagedSkills 'powerpoint-cli')
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
