#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook for PowerPointMcp: branch guard, Success-flag audit, COM leak audit,
    dynamic cast audit, Core interface completeness audit, build, targeted real-COM Core tests,
    MCP protocol tests, and a TODO/FIXME scan.

.DESCRIPTION
    Runs checks before allowing commits (ported and adapted from mcp-server-excel's
    scripts/pre-commit.ps1 — see .github/copilot-instructions.md for the full gate table):

    0. Process cleanup   - kills stale POWERPNT.exe / McpServer / CLI processes to avoid file locks
    1. Branch guard      - never commit directly to 'main'
    2. Success flag scan - flags any 'Success = true' followed nearby by a non-null ErrorMessage
                            assignment in touched Core files (Rule 1)
    2b. COM leak audit   - every 'dynamic' COM object in src/*.cs is released in a finally block
    2c. Dynamic cast audit - every '((dynamic))' cast has a justification comment
    2d. Core interface completeness - every implemented Core Commands method is declared on its
                            I*Commands interface (PowerPoint's action enums are generator-derived
                            FROM the interface, so an undeclared method is silently unreachable)
    3. Release build     - dotnet build Sbroenne.PowerPointMcp.slnx -c Release, 0 warnings/errors.
                            Fully SKIPPED for docs-only commits (.md, .changeset/, docs/, gh-pages/,
                            issue/PR templates) — there is no compiled surface to validate.
    4. Core tests        - surgical Feature=-filtered real-COM integration tests, scoped to the
                            Core domains touched by this commit. Skipped if no Core .cs changes,
                            and fully skipped (block not entered) for docs-only commits.
    5. MCP protocol tests - dotnet test tests\PowerPointMcp.McpServer.Tests. Fully skipped for
                            docs/changeset-only commits (including gh-pages/ website changes). For
                            docs/tooling changes that touch non-.cs files (scripts, workflows,
                            .gitignore) only the fast in-memory protocol tests run — the
                            PowerPoint-dependent COM session-lifecycle tests
                            (RequiresPowerPoint=true) are excluded. The full suite runs only when
                            compiled runtime code (src/tests *.cs, *.csproj, *.slnx,
                            Directory.Build/Packages, global.json) changes.
    6. TODO/FIXME/HACK scan - blocks unresolved markers in staged files

    NOTE: mcp-server-excel has additional release-packaging gates (NuGet pack, standalone ZIP,
    VS Code extension, MCPB bundle, Agent Skills ZIP) that do not yet apply here because
    PowerPoint's CLI/packaging surfaces are still maturing. Add those gates here as the
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

# Determine whether this commit touches actual code (as opposed to docs/changeset-only
# changes). The Release build, Core tests and MCP protocol test suite all exercise
# compiled binaries or launch real processes and are unnecessary for pure documentation
# changes, including edits to the gh-pages documentation website (its MkDocs config,
# hooks, templates and image assets are all part of the docs site, not the shipped
# product).
$docOnlyPattern = '(\.md$)|(^\.changeset/)|(^docs/)|(^gh-pages/)|(^\.github/(ISSUE_TEMPLATE|PULL_REQUEST_TEMPLATE))'
$allStagedFilesForGate = git diff --cached --name-only 2>&1 | Where-Object { $_ }
$codeChangedFilesForGate = $allStagedFilesForGate | Where-Object { $_ -notmatch $docOnlyPattern }
$hasCodeChanges = @($codeChangedFilesForGate).Count -gt 0

# Distinguish compiled runtime-code changes (which can alter MCP/COM behavior and therefore
# warrant the real-COM session-lifecycle tests) from docs/tooling changes (gh-pages, scripts,
# workflows, .gitignore) which cannot. The PowerPoint-dependent MCP tests are marked with the
# [Trait("RequiresPowerPoint","true")] attribute and are excluded unless runtime code changed.
$runtimeCodePattern = '(^src[/\\].*\.cs$)|(^tests[/\\].*\.cs$)|(\.csproj$)|(\.slnx$)|(^Directory\.(Build|Packages)\.)|(^global\.json$)|(^NuGet\.Config$)'
$runtimeCodeChanged = @($allStagedFilesForGate | Where-Object { $_ -match $runtimeCodePattern }).Count -gt 0

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

try {
    $successFlagScript = Join-Path $rootDir "scripts\check-success-flag.ps1"
    & $successFlagScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "BLOCKED: Success flag violations detected!" -ForegroundColor Red
        Write-Host "   Rule 1: Success==true must imply ErrorMessage==null. Fix before committing." -ForegroundColor Yellow
        exit 1
    }

    Write-Host "Success flag check passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running success flag check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- 2b. COM object leak audit (every 'dynamic' COM object released in a finally block) -----
if (-not $hasCodeChanges) {
    Write-Step "Skipping COM leak check (no code changes detected - docs/changeset only)"
}
else {
    Write-Step "Checking for COM object leaks..."

    try {
        $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
        & $leakCheckScript

        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "BLOCKED: COM object leaks detected! Fix them before committing." -ForegroundColor Red
            exit 1
        }

        Write-Host "COM leak check passed" -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# --- 2c. Dynamic cast audit (every '((dynamic))' cast has a justification comment) ----------
if (-not $hasCodeChanges) {
    Write-Step "Skipping dynamic cast audit (no code changes detected - docs/changeset only)"
}
else {
    Write-Step "Checking dynamic cast justifications..."

    try {
        $dynamicCastScript = Join-Path $rootDir "scripts\check-dynamic-casts.ps1"
        & $dynamicCastScript

        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "BLOCKED: Undocumented dynamic casts detected! Add a justification comment before committing." -ForegroundColor Red
            exit 1
        }

        Write-Host "Dynamic cast audit passed" -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "Error running dynamic cast audit: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# --- 2d. Core interface completeness audit (every implemented method is interface-declared) --
if (-not $hasCodeChanges) {
    Write-Step "Skipping Core interface completeness audit (no code changes detected - docs/changeset only)"
}
else {
    Write-Step "Checking Core interface completeness..."

    try {
        $interfaceScript = Join-Path $rootDir "scripts\check-core-interface-completeness.ps1"
        & $interfaceScript

        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "BLOCKED: Core Commands methods not declared on their interface detected!" -ForegroundColor Red
            Write-Host "   PowerPoint's action enums are generated FROM the I*Commands interface -" -ForegroundColor Yellow
            Write-Host "   an implemented-but-undeclared method is silently unreachable via MCP/CLI." -ForegroundColor Yellow
            exit 1
        }

        Write-Host "Core interface completeness audit passed" -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "Error running Core interface completeness audit: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# --- 3. Release build ------------------------------------------------------------------------
# Skipped entirely for docs-only commits (including gh-pages) - there is no compiled
# surface to validate, so the build gate adds minutes of cost for zero value.
if (-not $hasCodeChanges) {
    Write-Step "Skipping Release build (no code changes detected - docs/changeset only)"
}
else {
    Write-Step "Building Release solution..."

    $slnPath = Join-Path $rootDir "Sbroenne.PowerPointMcp.slnx"
    & dotnet build $slnPath -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "BLOCKED: Release build failed." -ForegroundColor Red
        exit 1
    }

    Write-Host "Release build passed (0 warnings, 0 errors expected)" -ForegroundColor Green
}

# --- 4. Surgical Core tests, scoped to touched domains ---------------------------------------
# Skipped entirely for docs-only commits (including gh-pages) - real-COM tests only make
# sense when Core source files actually changed.
if (-not $hasCodeChanges) {
    Write-Step "Skipping Core tests (no code changes detected - docs/changeset only)"
}
else {
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
}

# --- 5. MCP protocol tests (fast, in-memory transport) ---------------------------------------
if (-not $hasCodeChanges) {
    Write-Step "Skipping MCP Server tests (no code changes detected - docs/changeset only)"
}
elseif ($runtimeCodeChanged) {
    Write-Step "Running full MCP Server test suite (runtime code changed - includes real-COM session tests)..."

    & dotnet test (Join-Path $rootDir "tests\PowerPointMcp.McpServer.Tests") --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "BLOCKED: MCP Server tests failed." -ForegroundColor Red
        exit 1
    }

    Write-Host "MCP Server tests passed" -ForegroundColor Green
}
else {
    Write-Step "Running MCP Server protocol tests only (docs/tooling change - skipping PowerPoint-dependent COM session tests)..."

    & dotnet test (Join-Path $rootDir "tests\PowerPointMcp.McpServer.Tests") --filter "RequiresPowerPoint!=true" --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "BLOCKED: MCP Server protocol tests failed." -ForegroundColor Red
        exit 1
    }

    Write-Host "MCP Server protocol tests passed (COM session tests skipped for docs/tooling change)" -ForegroundColor Green
}

# --- 6. TODO/FIXME/HACK scan ------------------------------------------------------------------
Write-Step "Scanning staged files for TODO/FIXME/HACK markers..."

$allStagedFiles = git diff --cached --name-only --diff-filter=ACM
$markerHits = @()

foreach ($file in $allStagedFiles) {
    $fullPath = Join-Path $rootDir $file
    if (-not (Test-Path $fullPath)) { continue }
    if ($file -notmatch "\.(cs|md|ps1)$") { continue }

    # The pre-commit script itself necessarily contains the marker literals in its own scan
    # pattern and messages; skip it to avoid self-matching false positives.
    if ($file -match "scripts[/\\]pre-commit\.ps1$") { continue }

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
