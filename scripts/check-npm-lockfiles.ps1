#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    Rejects fixed HTTP download URLs in npm lockfiles without exposing their contents.
.DESCRIPTION
    Checks tracked package-lock.json and npm-shrinkwrap.json files, including nested
    projects. Use -Staged in the commit hook to check the index rather than working files.
#>
param(
    [string[]]$LockfilePath,
    [switch]$Staged,
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"

function Test-DownloadUrl {
    param($Value)

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            if ($key -eq "resolved" -and $Value[$key] -is [string] -and
                $Value[$key] -match '^(https?:)?//') {
                return $true
            }
            if (Test-DownloadUrl $Value[$key]) { return $true }
        }
    } elseif ($Value -is [array]) {
        foreach ($item in $Value) {
            if (Test-DownloadUrl $item) { return $true }
        }
    }
    return $false
}

if ($Staged -and $LockfilePath) {
    throw "Use -Staged with repository discovery, not -LockfilePath."
}

$explicitPaths = [bool]$LockfilePath
if (-not $explicitPaths) {
    $tracked = git -C $RepoRoot -c core.quotepath=false ls-files
    if ($LASTEXITCODE -ne 0) { throw "BLOCKED: Could not discover tracked npm lockfiles." }
    $LockfilePath = @($tracked | Where-Object { $_ -match '(^|/)(package-lock|npm-shrinkwrap)\.json$' })
}

foreach ($path in $LockfilePath) {
    if ($Staged) {
        $content = (git -C $RepoRoot show ":$path") -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "BLOCKED: Could not read staged lockfile: $path" }
    } else {
        $file = if ($explicitPaths) { $path } else { Join-Path $RepoRoot $path }
        $content = Get-Content -LiteralPath $file -Raw
    }

    try {
        $lockfile = ConvertFrom-Json -InputObject $content -AsHashtable
    } catch [System.ArgumentException] {
        throw "BLOCKED: Invalid JSON in npm lockfile: $path"
    }
    if ($lockfile -isnot [System.Collections.IDictionary] -or
        $lockfile.lockfileVersion -notin @(1, 2, 3)) {
        throw "BLOCKED: Unsupported npm lockfile format: $path"
    }
    if (Test-DownloadUrl $lockfile) {
        throw "BLOCKED: Fixed download URLs in $path. Run npm install --package-lock-only --ignore-scripts in that project, then stage the regenerated lockfile."
    }
}

Write-Host "Npm lockfile portability check passed ($(@($LockfilePath).Count) files)." -ForegroundColor Green
exit 0
