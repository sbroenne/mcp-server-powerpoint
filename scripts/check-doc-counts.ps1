#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Updates or validates advertised tool and operation counts against generated code metadata.

.DESCRIPTION
    Counts are derived from the generated skill manifest plus the hand-written
    PresentationToolAction enum. The result is cross-checked against the expected MCP
    protocol tool names, then written once to doc-counts.json and managed documentation
    claims. CI validates structure on pull requests; the main-branch workflow runs -Update.

.PARAMETER RepoRoot
    Repository root containing source code and generated build output.

.PARAMETER DocsRoot
    Root containing documents to update or validate. Defaults to RepoRoot.

.PARAMETER SkipBuild
    Use the current Release build output instead of refreshing it first.

.PARAMETER Update
    Update managed advertised totals and per-domain operation counts.

.PARAMETER AllowStaleAdvertisedCounts
    Allow numeric claims to differ while still checking their presence and count structure.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$DocsRoot,
    [switch]$SkipBuild,
    [switch]$Update,
    [switch]$AllowStaleAdvertisedCounts
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($DocsRoot)) {
    $DocsRoot = $RepoRoot
}
else {
    $DocsRoot = (Resolve-Path -LiteralPath $DocsRoot).Path
}
if ($Update -and $AllowStaleAdvertisedCounts) {
    throw '-Update and -AllowStaleAdvertisedCounts cannot be used together.'
}

if (-not $SkipBuild) {
    & dotnet build (Join-Path $RepoRoot 'Sbroenne.PowerPointMcp.slnx') -c Release --no-restore -p:NuGetAudit=false --verbosity minimal
    if ($LASTEXITCODE -ne 0) {
        throw 'Release build failed. Run dotnet restore, then retry.'
    }
}

$script:errors = [System.Collections.Generic.List[string]]::new()
$script:updatedFiles = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

function Add-Failure([string]$Message) {
    $script:errors.Add($Message)
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
$manifestOps = 0
$script:domainOpCounts = @{ presentation = $presentationOps }
$canonicalToolNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::Ordinal)
[void]$canonicalToolNames.Add('presentation')
foreach ($command in $manifest.commands) {
    $actionCount = @($command.actions).Count
    if ($actionCount -eq 0) {
        throw "Generated manifest command '$($command.name)' has no actions."
    }
    $manifestOps += $actionCount
    $script:domainOpCounts[[string]$command.name] = $actionCount
    [void]$canonicalToolNames.Add([string]$command.name)
}

$canonicalTools = $manifestTools + 1
$canonicalOperations = $manifestOps + $presentationOps
$canonicalDomains = $canonicalTools

# The generated manifest supplies every generated tool name; presentation is the
# hand-written tool. The protocol test verifies that this expected set is the live tools/list surface.
$protocolTestsPath = Join-Path $RepoRoot 'tests\PowerPointMcp.McpServer.Tests\Integration\McpProtocolTests.cs'
$protocolTestsContent = Get-Content -LiteralPath $protocolTestsPath -Raw
$expectedToolsMatch = [regex]::Match(
    $protocolTestsContent,
    '(?s)ExpectedToolNames\s*=\s*\[(?<body>.*?)\];')
if (-not $expectedToolsMatch.Success) {
    throw 'Could not locate McpProtocolTests.ExpectedToolNames.'
}
$expectedToolNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::Ordinal)
foreach ($match in [regex]::Matches($expectedToolsMatch.Groups['body'].Value, '"(?<name>[a-z]+)"')) {
    [void]$expectedToolNames.Add($match.Groups['name'].Value)
}
if (-not $canonicalToolNames.SetEquals($expectedToolNames)) {
    Add-Failure 'Generated manifest tool names do not match McpProtocolTests.ExpectedToolNames (the protocol test checks the live tools/list surface).'
}

$presentationToolsPath = Join-Path $RepoRoot 'src\PowerPointMcp.McpServer\Tools\PresentationTools.cs'
$presentationToolsContent = Get-Content -LiteralPath $presentationToolsPath -Raw
$handWrittenToolNames = @([regex]::Matches(
    $presentationToolsContent,
    'McpServerTool\s*\(\s*Name\s*=\s*"(?<name>[^"]+)"') |
    ForEach-Object { $_.Groups['name'].Value })
if (@($handWrittenToolNames).Count -ne 1 -or $handWrittenToolNames[0] -cne 'presentation') {
    Add-Failure 'Expected exactly one hand-written MCP tool named presentation in PresentationTools.cs.'
}

Write-Host "Canonical (from code): $canonicalTools tools, $canonicalOperations operations, $canonicalDomains domains" -ForegroundColor Cyan
Write-Host "  generated manifest: $manifestTools tools / $manifestOps operations; hand-written presentation: $presentationOps operations; protocol surface: $($expectedToolNames.Count) tools" -ForegroundColor DarkGray

$script:counts = @{
    t = $canonicalTools
    o = $canonicalOperations
    d = $canonicalDomains
    m = $manifestTools
}

function Get-DocumentPath([string]$RelativePath) {
    $path = Join-Path $DocsRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Add-Failure "Expected document not found: $RelativePath"
        return $null
    }
    return $path
}

function Save-Document([string]$RelativePath, [string]$Path, [string]$Content) {
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
    [void]$script:updatedFiles.Add($RelativePath)
}

function Update-CountPattern(
    [string]$RelativePath,
    [string]$Pattern,
    [string[]]$Groups
) {
    $path = Get-DocumentPath $RelativePath
    if (-not $path) { return }

    $content = Get-Content -LiteralPath $path -Raw
    $script:matchCount = 0
    $updated = [regex]::Replace($content, $Pattern, {
        param($match)
        $script:matchCount++
        $value = $match.Value
        foreach ($groupName in ($Groups | Sort-Object { $match.Groups[$_].Index } -Descending)) {
            $group = $match.Groups[$groupName]
            if (-not $group.Success) { continue }
            $expected = [string]$script:counts[$groupName]
            if ($group.Value -cne $expected) {
                if ($Update) {
                    $offset = $group.Index - $match.Index
                    $value = $value.Remove($offset, $group.Length).Insert($offset, $expected)
                }
                elseif (-not $AllowStaleAdvertisedCounts) {
                    Add-Failure ("{0}: {1} count is {2} but should be {3}." -f $RelativePath, $groupName, $group.Value, $expected)
                }
            }
        }
        $value
    })
    $matchCount = $script:matchCount
    $script:matchCount = 0
    if ($matchCount -eq 0) {
        Add-Failure "${RelativePath}: expected count pattern not found: /$Pattern/"
    }
    elseif ($Update -and $updated -cne $content) {
        Save-Document $RelativePath $path $updated
    }
}

function Update-DomainCounts([string]$RelativePath, [string]$Pattern) {
    $path = Get-DocumentPath $RelativePath
    if (-not $path) { return }

    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    $content = Get-Content -LiteralPath $path -Raw
    $updated = [regex]::Replace($content, $Pattern, {
        param($match)
        $name = $match.Groups['name'].Value -replace '[^A-Za-z]', ''
        if ($name -ieq 'customshows') { $name = 'customshow' }
        if (-not $script:domainOpCounts.ContainsKey($name)) { return $match.Value }

        [void]$seen.Add($name)
        $countGroup = $match.Groups['n']
        $expected = [string]$script:domainOpCounts[$name]
        if ($countGroup.Value -cne $expected) {
            if ($Update) {
                $offset = $countGroup.Index - $match.Index
                return $match.Value.Remove($offset, $countGroup.Length).Insert($offset, $expected)
            }
            if (-not $AllowStaleAdvertisedCounts) {
                Add-Failure ("{0}: {1} operation count is {2} but should be {3}." -f $RelativePath, $match.Groups['name'].Value, $countGroup.Value, $expected)
            }
        }
        $match.Value
    })

    $missing = @($script:domainOpCounts.Keys | Where-Object { -not $seen.Contains($_) })
    if ($missing.Count -gt 0) {
        Add-Failure "${RelativePath}: missing domain count entries for $($missing -join ', ')."
    }
    if ($Update -and $updated -cne $content) {
        Save-Document $RelativePath $path $updated
    }
}

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

Update-DomainCounts 'README.md' '\*\*(?<name>[A-Za-z ]+)\*\* \((?<n>\d+) ops?\)'
Update-DomainCounts 'src\PowerPointMcp.McpServer\README.md' '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|'
Update-DomainCounts 'gh-pages\docs\features.md' '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|'
Update-DomainCounts 'gh-pages\docs\features.md' '### `(?<name>[a-z]+)` tool \((?<n>\d+) operations\)'

$docCountsPath = Join-Path $DocsRoot 'doc-counts.json'
$docCountsJson = ([ordered]@{
    tools = $canonicalTools
    operations = $canonicalOperations
    domains = $canonicalDomains
} | ConvertTo-Json) + "`n"
if ($Update) {
    $current = if (Test-Path -LiteralPath $docCountsPath) {
        Get-Content -LiteralPath $docCountsPath -Raw
    }
    else { $null }
    if ($current -cne $docCountsJson) {
        [System.IO.File]::WriteAllText($docCountsPath, $docCountsJson, [System.Text.UTF8Encoding]::new($false))
        [void]$script:updatedFiles.Add('doc-counts.json')
    }
}
elseif (-not (Test-Path -LiteralPath $docCountsPath -PathType Leaf)) {
    Add-Failure 'doc-counts.json not found. Run this script with -Update to generate the canonical counts file.'
}
else {
    try {
        $storedCounts = Get-Content -LiteralPath $docCountsPath -Raw | ConvertFrom-Json
        $countKeys = @{
            tools = 't'
            operations = 'o'
            domains = 'd'
        }
        foreach ($key in $countKeys.Keys) {
            if ($null -eq $storedCounts.$key) {
                Add-Failure "doc-counts.json is missing the '$key' count."
            }
            elseif ([int]$storedCounts.$key -ne [int]$script:counts[$countKeys[$key]]) {
                if (-not $AllowStaleAdvertisedCounts) {
                    Add-Failure ("doc-counts.json: {0} count is {1} but should be {2}." -f $key, $storedCounts.$key, [int]$script:counts[$countKeys[$key]])
                }
            }
        }
    }
    catch {
        Add-Failure "doc-counts.json is malformed: $($_.Exception.Message)"
    }
}

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Documentation count validation FAILED ($($errors.Count) issue(s)):" -ForegroundColor Red
    foreach ($errorMessage in $errors) {
        Write-Host "  - $errorMessage" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Canonical counts are derived from code: $canonicalTools tools / $canonicalOperations operations." -ForegroundColor Yellow
    exit 1
}

if ($Update) {
    Write-Host "Documentation counts generated - $canonicalTools tools / $canonicalOperations operations ($($updatedFiles.Count) file(s) changed)" -ForegroundColor Green
}
elseif ($AllowStaleAdvertisedCounts) {
    Write-Host 'Documentation count structure passed; advertised totals may be behind main.' -ForegroundColor Green
}
else {
    Write-Host "Documentation count validation passed - all docs report $canonicalTools tools / $canonicalOperations operations" -ForegroundColor Green
}
exit 0
