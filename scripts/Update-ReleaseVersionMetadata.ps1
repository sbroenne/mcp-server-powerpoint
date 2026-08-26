<#
.SYNOPSIS
    Synchronizes persistent release versions across every distribution manifest.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

function Get-RequiredJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release metadata file was not found: $Path"
    }

    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Set-RootJsonVersion([string]$Path) {
    $json = Get-RequiredJson $Path
    if ($null -eq $json.PSObject.Properties['version']) {
        throw "Release metadata must contain a top-level 'version' property: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $pattern = '(?m)^(\s*"version"\s*:\s*")[^"]+(")'
    if ([regex]::Matches($content, $pattern).Count -lt 1) {
        throw "Release metadata version could not be located in: $Path"
    }

    $content = [regex]::new($pattern).Replace($content, "`${1}$Version`${2}", 1)
    [System.IO.File]::WriteAllText($Path, $content, [System.Text.UTF8Encoding]::new($false))
}

function Set-PackageLockVersion([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release metadata file was not found: $Path"
    }

    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    if (-not $json.ContainsKey('version') -or
        $null -eq $json.packages[''] -or
        -not $json.packages[''].ContainsKey('version')) {
        throw "Package lock must contain top-level and root-package versions: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $pattern = '(?m)^(\s*"version"\s*:\s*")[^"]+(")'
    if ([regex]::Matches($content, $pattern).Count -lt 2) {
        throw "Package lock versions could not be located in: $Path"
    }

    $content = [regex]::new($pattern).Replace($content, "`${1}$Version`${2}", 2)
    [System.IO.File]::WriteAllText($Path, $content, [System.Text.UTF8Encoding]::new($false))
}

function Set-ProjectVersions([string]$Path) {
    $content = Get-Content -LiteralPath $Path -Raw
    $replacements = [ordered]@{
        '<Version>[^<]+</Version>' = "<Version>$Version</Version>"
        '<AssemblyVersion>[^<]+</AssemblyVersion>' = "<AssemblyVersion>$Version.0</AssemblyVersion>"
        '<FileVersion>[^<]+</FileVersion>' = "<FileVersion>$Version.0</FileVersion>"
    }

    foreach ($replacement in $replacements.GetEnumerator()) {
        if ([regex]::Matches($content, $replacement.Key).Count -ne 1) {
            throw "Expected exactly one '$($replacement.Key)' value in $Path."
        }

        $content = [regex]::Replace($content, $replacement.Key, $replacement.Value)
    }

    [System.IO.File]::WriteAllText($Path, $content, [System.Text.UTF8Encoding]::new($false))
}

Set-RootJsonVersion (Join-Path $RepoRoot 'package.json')
Set-PackageLockVersion (Join-Path $RepoRoot 'package-lock.json')
Set-ProjectVersions (Join-Path $RepoRoot 'Directory.Build.props')
Set-RootJsonVersion (Join-Path $RepoRoot 'mcpb' 'manifest.json')
Set-RootJsonVersion (Join-Path $RepoRoot 'vscode-extension' 'package.json')
Set-PackageLockVersion (Join-Path $RepoRoot 'vscode-extension' 'package-lock.json')

& (Join-Path $PSScriptRoot 'Update-McpRegistryMetadata.ps1') `
    -ServerJsonPath (Join-Path $RepoRoot 'src' 'PowerPointMcp.McpServer' '.mcp' 'server.json') `
    -Version $Version

Write-Output "Synchronized persistent release metadata to version $Version."
