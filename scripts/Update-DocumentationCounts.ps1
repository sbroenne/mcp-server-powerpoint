#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Updates user-facing tool and operation counts from generated code metadata.

.DESCRIPTION
    Derives the canonical MCP surface from the generated skill manifest plus the
    hand-written presentation action enum, verifies it against the MCP tool
    registrations, and updates the release documentation in place.

.PARAMETER RepoRoot
    Repository root containing source code and generated build output.

.PARAMETER DocsRoot
    Root containing the documents to update. Defaults to RepoRoot.

.PARAMETER SkipBuild
    Uses the current Release build output instead of refreshing it first.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),

    [string]$DocsRoot,

    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($DocsRoot)) {
    $DocsRoot = $RepoRoot
}
else {
    $DocsRoot = (Resolve-Path -LiteralPath $DocsRoot).Path
}

if (-not $SkipBuild) {
    & dotnet build (Join-Path $RepoRoot 'Sbroenne.PowerPointMcp.slnx') -c Release --no-restore -p:NuGetAudit=false --verbosity minimal
    if ($LASTEXITCODE -ne 0) {
        throw 'Release build failed. Run dotnet restore, then retry.'
    }
}

$manifestFile = Get-ChildItem -Path (Join-Path $RepoRoot 'src\PowerPointMcp.Core\obj') -Recurse -Filter '_SkillManifest.g.cs' -ErrorAction SilentlyContinue |
    Sort-Object { $_.FullName -notmatch 'GeneratedFiles' } |
    Select-Object -First 1
if (-not $manifestFile) {
    throw 'Could not find generated _SkillManifest.g.cs. Complete a Release build first.'
}

$manifestContent = Get-Content -LiteralPath $manifestFile.FullName -Raw
$startMarker = 'public const string Json = @"'
$startIdx = $manifestContent.IndexOf($startMarker)
$endIdx = $manifestContent.LastIndexOf('";')
if ($startIdx -lt 0 -or $endIdx -le $startIdx) {
    throw "Could not extract JSON from $($manifestFile.FullName)."
}
$startIdx += $startMarker.Length
$manifest = $manifestContent.Substring($startIdx, $endIdx - $startIdx).Replace('""', '"') |
    ConvertFrom-Json

$presentationActionPath = Join-Path $RepoRoot 'src\PowerPointMcp.Core\Presentation\PresentationToolAction.cs'
$presentationActionContent = Get-Content -LiteralPath $presentationActionPath -Raw
$enumMatch = [regex]::Match(
    $presentationActionContent,
    '(?s)public enum PresentationToolAction\s*\{(?<body>.*?)\n\}')
if (-not $enumMatch.Success) {
    throw 'Could not locate the PresentationToolAction enum body.'
}
$presentationOps = [regex]::Matches(
    $enumMatch.Groups['body'].Value,
    'JsonStringEnumMemberName').Count
if ($presentationOps -eq 0) {
    throw 'PresentationToolAction enum parsed to zero operations.'
}

$manifestTools = @($manifest.commands).Count
$manifestOps = ($manifest.commands |
    ForEach-Object { @($_.actions).Count } |
    Measure-Object -Sum).Sum
$canonicalTools = $manifestTools + 1
$canonicalOperations = $manifestOps + $presentationOps

$protocolTestsPath = Join-Path $RepoRoot 'tests\PowerPointMcp.McpServer.Tests\Integration\McpProtocolTests.cs'
$protocolTestsContent = Get-Content -LiteralPath $protocolTestsPath -Raw
$expectedToolsMatch = [regex]::Match(
    $protocolTestsContent,
    '(?s)ExpectedToolNames\s*=\s*\[(?<body>.*?)\];')
if (-not $expectedToolsMatch.Success) {
    throw 'Could not locate McpProtocolTests.ExpectedToolNames.'
}
$expectedToolNames = [System.Collections.Generic.HashSet[string]]::new()
foreach ($match in [regex]::Matches($expectedToolsMatch.Groups['body'].Value, '"(?<name>[a-z]+)"')) {
    [void]$expectedToolNames.Add($match.Groups['name'].Value)
}
$canonicalToolNames = [System.Collections.Generic.HashSet[string]]::new()
[void]$canonicalToolNames.Add('presentation')
foreach ($command in $manifest.commands) {
    [void]$canonicalToolNames.Add($command.name)
}
if (-not $canonicalToolNames.SetEquals($expectedToolNames)) {
    throw 'Generated manifest tool names do not match McpProtocolTests.ExpectedToolNames.'
}

$counts = @{
    t = $canonicalTools
    o = $canonicalOperations
    d = $canonicalTools
    m = $manifestTools
}

function Update-CountPattern(
    [string]$RelativePath,
    [string]$Pattern,
    [string[]]$Groups
) {
    $path = Join-Path $DocsRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Expected document not found: $RelativePath"
    }

    $content = Get-Content -LiteralPath $path -Raw
    $matches = [System.Collections.Generic.List[object]]::new()
    $updated = [regex]::Replace($content, $Pattern, {
        param($match)
        $matches.Add($match)
        $value = $match.Value
        foreach ($groupName in ($Groups | Sort-Object { $match.Groups[$_].Index } -Descending)) {
            $group = $match.Groups[$groupName]
            if (-not $group.Success) {
                continue
            }
            $offset = $group.Index - $match.Index
            $replacement = [string]$counts[$groupName]
            $value = $value.Remove($offset, $group.Length).Insert($offset, $replacement)
        }
        $value
    })
    if ($matches.Count -eq 0) {
        throw "${RelativePath}: expected count pattern not found: /$Pattern/"
    }
    if ($updated -cne $content) {
        Set-Content -LiteralPath $path -Value $updated -NoNewline
        $script:updatedFiles.Add($RelativePath)
    }
}

function Update-DomainCounts([string]$RelativePath, [string]$Pattern) {
    $path = Join-Path $DocsRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Expected document not found: $RelativePath"
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    $content = Get-Content -LiteralPath $path -Raw
    $updated = [regex]::Replace($content, $Pattern, {
        param($match)
        $name = $match.Groups['name'].Value -replace '[^A-Za-z]', ''
        if ($name -ieq 'customshows') {
            $name = 'customshow'
        }
        if (-not $script:domainOpCounts.ContainsKey($name)) {
            return $match.Value
        }

        [void]$seen.Add($name)
        $countGroup = $match.Groups['n']
        $offset = $countGroup.Index - $match.Index
        $match.Value.Remove($offset, $countGroup.Length).Insert(
            $offset,
            [string]$script:domainOpCounts[$name])
    })

    $missing = @($script:domainOpCounts.Keys | Where-Object { -not $seen.Contains($_) })
    if ($missing.Count -gt 0) {
        throw "${RelativePath}: missing domain count entries for $($missing -join ', ')."
    }
    if ($updated -cne $content) {
        Set-Content -LiteralPath $path -Value $updated -NoNewline
        $script:updatedFiles.Add($RelativePath)
    }
}

$script:updatedFiles = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

$headlineChecks = @(
    @{ File = 'README.md'; Pattern = '(?<t>\d+) MCP tools with (?<o>\d+) operations across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'src\PowerPointMcp.McpServer\README.md'; Pattern = '(?<t>\d+) tools with (?<o>\d+) operations across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'mcpb\README.md'; Pattern = 'tools \((?<o>\d+) operations across (?<d>\d+) domains\)'; Groups = @('o', 'd') }
    @{ File = 'mcpb\manifest.json'; Pattern = '(?<t>\d+) tools \((?<o>\d+) operations across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'gh-pages\docs\index.md'; Pattern = 'all (?<t>\d+) tools \((?<o>\d+) operations\) across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'gh-pages\docs\installation.md'; Pattern = 'all (?<t>\d+) tools \((?<o>\d+) operations\) across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'gh-pages\docs\features.md'; Pattern = '(?<t>\d+) MCP tools with (?<o>\d+) operations across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'gh-pages\docs\mcp-server.md'; Pattern = '(?<t>\d+) tools with (?<o>\d+) operations across (?<d>\d+) domains'; Groups = @('t', 'o', 'd') }
    @{ File = 'skills\CLAUDE.md'; Pattern = '(?<t>\d+) MCP tools across (?<d>\d+) domains'; Groups = @('t', 'd') }
    @{ File = 'skills\powerpoint-mcp\SKILL.md'; Pattern = 'Provides (?<t>\d+) PowerPoint MCP tools \(one presentation tool \+ (?<m>\d+) domain action-dispatch tools\)'; Groups = @('t', 'm') }
    @{ File = 'skills\shared\behavioral-rules.md'; Pattern = '(?<t>\d+) PowerPoint MCP tools across (?<d>\d+) domains'; Groups = @('t', 'd') }
    @{ File = 'skills\shared\behavioral-rules.md'; Pattern = 'All (?<t>\d+) MCP tools are action-dispatch tools'; Groups = @('t') }
    @{ File = 'skills\shared\behavioral-rules.md'; Pattern = 'The other (?<m>\d+) domain tools'; Groups = @('m') }
    @{ File = 'skills\shared\workflows.md'; Pattern = 'All (?<t>\d+) tools and (?<o>\d+) operations'; Groups = @('t', 'o') }
)
foreach ($check in $headlineChecks) {
    Update-CountPattern $check.File $check.Pattern $check.Groups
}

$script:domainOpCounts = @{ presentation = $presentationOps }
foreach ($command in $manifest.commands) {
    $script:domainOpCounts[$command.name] = @($command.actions).Count
}

Update-DomainCounts 'README.md' '\*\*(?<name>[A-Za-z ]+)\*\* \((?<n>\d+) ops?\)'
Update-DomainCounts 'src\PowerPointMcp.McpServer\README.md' '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|'
Update-DomainCounts 'gh-pages\docs\features.md' '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|'
Update-DomainCounts 'gh-pages\docs\features.md' '### `(?<name>[a-z]+)` tool \((?<n>\d+) operations\)'

Write-Output "Canonical documentation counts: $canonicalTools tools, $canonicalOperations operations."
if ($updatedFiles.Count -eq 0) {
    Write-Output 'Documentation counts are already current.'
}
else {
    Write-Output "Updated $($updatedFiles.Count) document(s):"
    $updatedFiles | Sort-Object | ForEach-Object { Write-Output "  $_" }
}
