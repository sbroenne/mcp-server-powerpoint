#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates that the tool/operation counts advertised in user-facing docs match the
    authoritative counts derived from the code (generated skill manifest + the hand-written
    presentation tool's action enum).

.DESCRIPTION
    Ported from mcp-server-excel's scripts/check-doc-counts.ps1, adapted for PowerPointMcp's
    architecture (see .github/instructions/mcp-server-guide.instructions.md and
    IPresentationCommands.cs remarks):

    - Excel derives its canonical count from a generated manifest PLUS a hand-written
      FileAction enum (the session/file tool), MINUS a CLI-only 'diag' category.
    - PowerPointMcp has no 'diag' category to subtract, but has the equivalent hand-written
      'presentation' tool (PresentationToolAction enum) to ADD, because IPresentationCommands
      deliberately carries no [ServiceCategory] attribute and is therefore absent from the
      generated manifest (mirrors Excel's IFileCommands exactly).

    THE PROBLEM IT PREVENTS
    -----------------------
    Several user-facing docs (README.md, gh-pages, MCP Server README, mcpb README, skills)
    hard-code "13 tools / 132 operations" from memory. That count silently drifted out of sync
    when Image gained set-crop/get-crop (132 -> 134) - the docs were never updated. This script
    computes the ONE canonical answer from code on every commit and fails if any doc disagrees.

    HOW THE CANONICAL NUMBERS ARE DERIVED
    --------------------------------------
    Authoritative source = the generated `_SkillManifest.g.cs` (produced by
    ServiceRegistryGenerator from the Core [ServiceCategory] interfaces). It reports every
    generator-driven domain (slide, shape, textframe, table, notes, layout, master, animation,
    image, chart, smartart, export) - 12 tools.

    canonical tools      = manifest.Commands.Count + 1              (+ presentation)
    canonical operations = sum(manifest.Commands[*].Actions.Count)
                           + PresentationToolAction enum member count

.NOTES
    Run after a Release build so the generated manifest is current.
    Exit code 0 = all counts consistent. Exit code 1 = a mismatch was found.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

$errors = [System.Collections.Generic.List[string]]::new()
function Add-Failure([string]$message) { $script:errors.Add($message) }

# ---------------------------------------------------------------------------
# 1. Parse the authoritative generated skill manifest
# ---------------------------------------------------------------------------
$manifestFile = Get-ChildItem -Path (Join-Path $rootDir "src\PowerPointMcp.Core\obj") -Recurse -Filter "_SkillManifest.g.cs" -ErrorAction SilentlyContinue |
    Sort-Object { $_.FullName -notmatch "GeneratedFiles" } |
    Select-Object -First 1

if (-not $manifestFile) {
    Write-Host "ERROR: Could not find generated _SkillManifest.g.cs. Run a Release build first." -ForegroundColor Red
    exit 1
}

$manifestContent = Get-Content $manifestFile.FullName -Raw
$startMarker = 'public const string Json = @"'
$startIdx = $manifestContent.IndexOf($startMarker)
$endIdx = $manifestContent.LastIndexOf('";')
if ($startIdx -lt 0 -or $endIdx -le $startIdx) {
    Write-Host "ERROR: Could not extract JSON from $($manifestFile.FullName)" -ForegroundColor Red
    exit 1
}
$startIdx += $startMarker.Length
$json = $manifestContent.Substring($startIdx, $endIdx - $startIdx).Replace('""', '"')
$manifest = $json | ConvertFrom-Json

$manifestTools = @($manifest.commands).Count
$manifestOps = ($manifest.commands | ForEach-Object { @($_.actions).Count } | Measure-Object -Sum).Sum

# ---------------------------------------------------------------------------
# 2. Compute the 'presentation' tool's operation count from ground truth
#    (hand-written - absent from the manifest by design, see IPresentationCommands.cs)
# ---------------------------------------------------------------------------
$presentationActionPath = Join-Path $rootDir "src\PowerPointMcp.Core\Presentation\PresentationToolAction.cs"
if (-not (Test-Path $presentationActionPath)) {
    Write-Host "ERROR: Could not locate PresentationToolAction.cs" -ForegroundColor Red
    exit 1
}
$presentationActionContent = Get-Content $presentationActionPath -Raw
$enumMatch = [regex]::Match($presentationActionContent, '(?s)public enum PresentationToolAction\s*\{(?<body>.*?)\n\}')
if (-not $enumMatch.Success) {
    Write-Host "ERROR: Could not locate the PresentationToolAction enum body" -ForegroundColor Red
    exit 1
}
$presentationOps = ([regex]::Matches($enumMatch.Groups['body'].Value, 'JsonStringEnumMemberName')).Count
if ($presentationOps -eq 0) {
    Write-Host "ERROR: PresentationToolAction enum parsed to 0 operations - parsing bug." -ForegroundColor Red
    exit 1
}

$canonicalTools = $manifestTools + 1
$canonicalOps = $manifestOps + $presentationOps

# ---------------------------------------------------------------------------
# 3. Cross-check against the REAL MCP tool surface ([McpServerTool(Name=...)])
# ---------------------------------------------------------------------------
$mcpToolNames = [System.Collections.Generic.HashSet[string]]::new()
$mcpDir = Join-Path $rootDir "src\PowerPointMcp.McpServer"
if (Test-Path $mcpDir) {
    Get-ChildItem -Path $mcpDir -Recurse -Filter "*.cs" -ErrorAction SilentlyContinue | ForEach-Object {
        $c = Get-Content $_.FullName -Raw
        foreach ($m in [regex]::Matches($c, 'McpServerTool\s*\(\s*Name\s*=\s*"([^"]+)"')) {
            [void]$mcpToolNames.Add($m.Groups[1].Value)
        }
    }
}

if ($mcpToolNames.Count -ne $canonicalTools) {
    Add-Failure ("MCP tool surface has {0} tools ([McpServerTool(Name=...)]) but the manifest-derived canonical tool count is {1}. If you added/removed a tool, update the docs; if the presentation tool's action count changed, this script picks it up automatically - re-run after a Release build." -f $mcpToolNames.Count, $canonicalTools)
}
if (-not $mcpToolNames.Contains('presentation')) {
    Add-Failure "No 'presentation' MCP tool found - the canonical count assumption (presentation adds $presentationOps ops) is broken."
}

Write-Host "Canonical (from code): $canonicalTools tools, $canonicalOps operations" -ForegroundColor Cyan
Write-Host "  manifest: $manifestTools tools / $manifestOps ops; + presentation($presentationOps); MCP tool surface: $($mcpToolNames.Count) tools" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 4. Validate headline claims across user-facing docs
# ---------------------------------------------------------------------------
# Each check: file + regex. Capture group 't' (optional) must equal canonicalTools,
# capture group 'o' (optional) must equal canonicalOps. A check that matches nothing fails
# (so a headline can't silently disappear or be reworded past the guard).
$checks = @(
    @{ File = "README.md";                                     Pattern = '(?<t>\d+) MCP tools with (?<o>\d+) operations' }
    @{ File = "src\PowerPointMcp.McpServer\README.md";          Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ File = "mcpb\README.md";                                 Pattern = 'tools \((?<o>\d+) operations across \d+ domains\)' }
    @{ File = "gh-pages\docs\index.md";                         Pattern = 'all (?<t>\d+) tools \((?<o>\d+) operations\) across \d+ domains' }
    @{ File = "gh-pages\docs\installation.md";                  Pattern = 'all (?<t>\d+) tools \((?<o>\d+) operations\) across \d+ domains' }
    @{ File = "gh-pages\docs\features.md";                      Pattern = '(?<t>\d+) MCP tools with (?<o>\d+) operations across \d+ domains' }
    @{ File = "gh-pages\docs\mcp-server.md";                    Pattern = '(?<t>\d+) tools with (?<o>\d+) operations across \d+ domains' }
    @{ File = "skills\shared\workflows.md";                     Pattern = 'All (?<t>\d+) tools and (?<o>\d+) operations' }
    @{ File = "skills\powerpoint-mcp\references\workflows.md";  Pattern = 'All (?<t>\d+) tools and (?<o>\d+) operations' }
)

foreach ($check in $checks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) {
        Add-Failure "Expected doc not found: $($check.File)"
        continue
    }
    $content = Get-Content $path -Raw
    $matches = [regex]::Matches($content, $check.Pattern)
    if ($matches.Count -eq 0) {
        Add-Failure "$($check.File): expected headline pattern not found (was it reworded or removed?): /$($check.Pattern)/"
        continue
    }
    foreach ($m in $matches) {
        if ($m.Groups['t'].Success -and [int]$m.Groups['t'].Value -ne $canonicalTools) {
            Add-Failure ("$($check.File): tool count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['t'].Value, $canonicalTools, $m.Value.Trim())
        }
        if ($m.Groups['o'].Success -and [int]$m.Groups['o'].Value -ne $canonicalOps) {
            Add-Failure ("$($check.File): operation count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['o'].Value, $canonicalOps, $m.Value.Trim())
        }
    }
}

# ---------------------------------------------------------------------------
# 5. Per-domain op counts: cross-check the manifest's per-command action count
#    against the two docs that enumerate every domain (README.md and the
#    MCP Server README's tool table). Catches drift like "image (5 ops)" after
#    a domain gains actions (e.g. Image's set-crop/get-crop).
# ---------------------------------------------------------------------------
$perDomainChecks = @(
    @{ File = "README.md";                              Pattern = '\*\*(?<name>[A-Za-z]+)\*\* \((?<n>\d+) ops?\)' }
    @{ File = "src\PowerPointMcp.McpServer\README.md";   Pattern = '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|' }
    @{ File = "gh-pages\docs\features.md";               Pattern = '`(?<name>[a-z]+)`\s*\|\s*(?<n>\d+)\s*\|' }
)

# name (case-insensitive, first-word-of-doc-label) -> actual action count from the manifest.
$domainOpCounts = @{ "presentation" = $presentationOps }
foreach ($cmd in $manifest.commands) {
    $domainOpCounts[$cmd.name] = @($cmd.actions).Count
}

foreach ($check in $perDomainChecks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) { continue }
    $content = Get-Content $path -Raw
    foreach ($m in [regex]::Matches($content, $check.Pattern)) {
        $label = $m.Groups['name'].Value.ToLowerInvariant()
        if (-not $domainOpCounts.ContainsKey($label)) { continue } # not a domain row (e.g. a table header)
        $claimed = [int]$m.Groups['n'].Value
        if ($claimed -ne $domainOpCounts[$label]) {
            Add-Failure ("$($check.File): '$label' claims $claimed ops but the manifest says $($domainOpCounts[$label]) -> `"$($m.Value.Trim())`"")
        }
    }
}

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Documentation count validation FAILED ($($errors.Count) issue(s)):" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host "  - $e" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Canonical counts are derived from code: $canonicalTools tools / $canonicalOps operations." -ForegroundColor Yellow
    Write-Host "Update the docs above to match, or if the surface genuinely changed, update the counts everywhere." -ForegroundColor Yellow
    exit 1
}

Write-Host "Documentation count validation passed - all docs report $canonicalTools tools / $canonicalOps operations" -ForegroundColor Green
exit 0
