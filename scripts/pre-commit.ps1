#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook for PowerPointMcp: branch guard, Success-flag audit, build, targeted
    real-COM Core tests, MCP protocol tests, and a TODO/FIXME scan.

.DESCRIPTION
    Runs checks before allowing commits (ported and adapted from mcp-server-excel's
    scripts/pre-commit.ps1 — see .github/copilot-instructions.md for the full gate table):

    0. Process cleanup   - kills stale POWERPNT.exe / McpServer / CLI processes to avoid file locks
    1. Branch guard      - never commit directly to 'main'
    2. Success flag scan - flags any 'Success = true' followed nearby by a non-null ErrorMessage
                            assignment in touched Core files (Rule 1)
    3. Release build     - dotnet build Sbroenne.PowerPointMcp.slnx -c Release, 0 warnings/errors
    4. Core tests        - surgical Feature=-filtered real-COM integration tests, scoped to the
                            Core domains touched by this commit (skipped if no Core changes)
    5. MCP protocol tests - dotnet test tests\PowerPointMcp.McpServer.Tests (in-memory transport,
                            fast, no PowerPoint launch required for most cases)
    6. TODO/FIXME/HACK scan - blocks unresolved markers in staged files

    NOTE: Unlike mcp-server-excel, this repo does not yet have companion scripts
    (check-com-leaks.ps1, audit-core-coverage.ps1, release-packaging checks, etc.) because the
    CLI/Generators/Service/packaging surfaces haven't been built yet. Add those gates here as the
    corresponding deliverables land — see .squad/decisions.md for the release-deliverables roadmap.

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Stop-DotNetBuildServers {
    dotnet build-server shutdown *> $null
}

# --- 1. Branch guard (never commit directly to main) ---------------------------------------
Write-Step "Checking current branch..."
$currentBranch = git branch --show-current

if ($currentBranch -eq "main") {
    Write-Host ""
    Write-Host "BLOCKED: Cannot commit directly to 'main' branch!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   All changes must go through a feature branch -> PR -> review -> merge." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   To fix:" -ForegroundColor Cyan
    Write-Host "   1. git stash                                    # Save your changes" -ForegroundColor White
    Write-Host "   2. git checkout -b squad/your-feature-name      # Create feature branch" -ForegroundColor White
    Write-Host "   3. git stash pop                                # Restore changes" -ForegroundColor White
    Write-Host "   4. git add <files>                              # Stage changes" -ForegroundColor White
    Write-Host "   5. git commit -m 'your message'                 # Commit to feature branch" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "Branch check passed - on '$currentBranch' (not main)" -ForegroundColor Green

# --- 0. Process cleanup (avoid file locks on Release binaries / open .pptx files) -----------
Write-Step "Killing stale PowerPoint and server processes..."

$killedProcesses = @()
foreach ($procName in @("POWERPNT", "Sbroenne.PowerPointMcp.McpServer", "Sbroenne.PowerPointMcp.CLI")) {
    $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        $killedProcesses += "$procName ($($procs.Count))"
    }
}

if ($killedProcesses.Count -gt 0) {
    Write-Host "   Killed: $($killedProcesses -join ', ')" -ForegroundColor Yellow
    Start-Sleep -Milliseconds 500
}
else {
    Write-Host "   No stale processes found" -ForegroundColor Gray
}

Stop-DotNetBuildServers
Write-Host "Process cleanup done" -ForegroundColor Green

# --- 2. Success flag audit (Rule 1: Success=true must never pair with ErrorMessage) ---------
Write-Step "Checking Success flag violations (Rule 1)..."

$stagedCsFiles = git diff --cached --name-only --diff-filter=ACM | Where-Object { $_ -like "src/*.cs" -or $_ -like "src\*.cs" }
$successFlagViolations = @()

foreach ($file in $stagedCsFiles) {
    $fullPath = Join-Path $rootDir $file
    if (-not (Test-Path $fullPath)) { continue }

    $lines = Get-Content $fullPath
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "Success\s*=\s*true") {
            # Look a few lines ahead within the same object initializer for a non-null ErrorMessage
            $window = $lines[$i..[Math]::Min($i + 5, $lines.Count - 1)] -join "`n"
            if ($window -match "ErrorMessage\s*=\s*[^n][^u][^l][^l]") {
                $successFlagViolations += "$file`:$($i + 1)"
            }
        }
    }
}

if ($successFlagViolations.Count -gt 0) {
    Write-Host ""
    Write-Host "BLOCKED: Possible Success=true + non-null ErrorMessage violation(s):" -ForegroundColor Red
    $successFlagViolations | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }
    Write-Host "   Rule 1: Success==true must imply ErrorMessage==null. Fix before committing." -ForegroundColor Yellow
    exit 1
}

Write-Host "Success flag check passed" -ForegroundColor Green

# --- 3. Release build ------------------------------------------------------------------------
Write-Step "Building Release solution..."

$slnPath = Join-Path $rootDir "Sbroenne.PowerPointMcp.slnx"
& dotnet build $slnPath -c Release --nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "BLOCKED: Release build failed." -ForegroundColor Red
    exit 1
}

Write-Host "Release build passed (0 warnings, 0 errors expected)" -ForegroundColor Green

# --- 4. Surgical Core tests, scoped to touched domains ---------------------------------------
Write-Step "Checking for touched Core domains..."

$touchedFeatures = @()
foreach ($file in $stagedCsFiles) {
    if ($file -match "PowerPointMcp\.Core[/\\](?<domain>[A-Za-z]+)[/\\]") {
        $touchedFeatures += $Matches.domain
    }
}
$touchedFeatures = $touchedFeatures | Select-Object -Unique

if ($touchedFeatures.Count -eq 0) {
    Write-Host "   No Core domain changes detected - skipping real-COM integration tests" -ForegroundColor Gray
}
else {
    foreach ($feature in $touchedFeatures) {
        Write-Step "Running real-COM tests for Feature=$feature (this launches PowerPoint)..."
        & dotnet test (Join-Path $rootDir "tests\PowerPointMcp.Core.Tests") --filter "Feature=$feature" --nologo
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "BLOCKED: Core tests failed for Feature=$feature." -ForegroundColor Red
            exit 1
        }
    }
    Write-Host "Core domain tests passed for: $($touchedFeatures -join ', ')" -ForegroundColor Green
}

# --- 5. MCP protocol tests (fast, in-memory transport) ---------------------------------------
Write-Step "Running MCP Server tests..."

& dotnet test (Join-Path $rootDir "tests\PowerPointMcp.McpServer.Tests") --nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "BLOCKED: MCP Server tests failed." -ForegroundColor Red
    exit 1
}

Write-Host "MCP Server tests passed" -ForegroundColor Green

# --- 6. TODO/FIXME/HACK scan ------------------------------------------------------------------
Write-Step "Scanning staged files for TODO/FIXME/HACK markers..."

$allStagedFiles = git diff --cached --name-only --diff-filter=ACM
$markerHits = @()

foreach ($file in $allStagedFiles) {
    $fullPath = Join-Path $rootDir $file
    if (-not (Test-Path $fullPath)) { continue }
    if ($file -notmatch "\.(cs|md|ps1)$") { continue }

    $matches = Select-String -Path $fullPath -Pattern "TODO|FIXME|HACK" -SimpleMatch -ErrorAction SilentlyContinue
    if ($matches) {
        $markerHits += $matches | ForEach-Object { "$file`:$($_.LineNumber)" }
    }
}

if ($markerHits.Count -gt 0) {
    Write-Host ""
    Write-Host "BLOCKED: Unresolved TODO/FIXME/HACK markers found:" -ForegroundColor Red
    $markerHits | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }
    Write-Host "   Resolve or explicitly defer (with a tracked issue link) before committing." -ForegroundColor Yellow
    exit 1
}

Write-Host "No unresolved TODO/FIXME/HACK markers" -ForegroundColor Green

# --- Done --------------------------------------------------------------------------------------
Write-Host ""
Write-Host "All pre-commit checks passed." -ForegroundColor Green
exit 0
