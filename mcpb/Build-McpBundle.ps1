<#
.SYNOPSIS
    Packages the direct-npx Windows MCP server configuration for Claude Desktop.
.DESCRIPTION
    Bundles metadata only. An available npx resolves the latest public
    PowerPointMcp package at launch.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = './artifacts'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'McpbPackaging.ps1')
. (Join-Path $root 'scripts\PackageHelpers.ps1')

if (-not $Version) {
    [xml]$props = Get-Content (Join-Path $root 'Directory.Build.props')
    $Version = $props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
}
if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$') {
    throw 'A valid package version is required.'
}

$output = [IO.Path]::GetFullPath($OutputDir, $PSScriptRoot)
Assert-PackageOutputPath -Path $output -RepoRoot $root
$stage = Join-Path ([IO.Path]::GetTempPath()) "PowerPointMcpMcpb-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $stage | Out-Null

try {
    $manifest = Get-Content (Join-Path $PSScriptRoot 'manifest.json') -Raw | ConvertFrom-Json
    $manifest.version = $Version
    $manifest | ConvertTo-Json -Depth 20 |
        Set-Content (Join-Path $stage 'manifest.json') -Encoding utf8

    foreach ($name in @('icon-512.png', 'README.md')) {
        Copy-Item (Join-Path $PSScriptRoot $name) $stage
    }
    foreach ($name in @('LICENSE', 'CHANGELOG.md')) {
        Copy-Item (Join-Path $root $name) $stage
    }

    $entries = @('manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md') |
        ForEach-Object { Join-Path $stage $_ }
    $archive = Join-Path $stage 'package.zip'
    Compress-Archive -LiteralPath $entries -DestinationPath $archive -CompressionLevel Optimal

    $expectedEntries = @('CHANGELOG.md', 'LICENSE', 'README.md', 'icon-512.png', 'manifest.json')
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $actualEntries = @($zip.Entries | ForEach-Object FullName | Sort-Object)
        if (Compare-Object $expectedEntries $actualEntries) {
            throw "MCPB content mismatch. Expected: $($expectedEntries -join ', '). Actual: $($actualEntries -join ', ')."
        }
    }
    finally {
        $zip.Dispose()
    }

    New-Item -ItemType Directory -Path $output -Force | Out-Null
    Install-PackageOutput `
        -Source $archive `
        -Destination (Join-Path $output "powerpoint-mcp-$Version.mcpb")
    Install-PackageOutput `
        -Source (Join-Path $stage 'manifest.json') `
        -Destination (Join-Path $output 'manifest.json')
}
finally {
    Remove-McpbStagingDirectory -Path $stage
}
