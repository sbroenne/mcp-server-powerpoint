[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ServerJsonPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$packageId = 'Sbroenne.PowerPointMcp.McpServer'

if (-not (Test-Path -LiteralPath $ServerJsonPath -PathType Leaf)) {
    throw "MCP Registry metadata file was not found: $ServerJsonPath"
}

$server = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
if ($null -eq $server.PSObject.Properties['version'] -or
    $null -eq $server.PSObject.Properties['packages']) {
    throw "MCP Registry metadata must contain top-level 'version' and 'packages' properties."
}

$packages = @($server.packages | Where-Object { $_.identifier -eq $packageId })
if ($packages.Count -ne 1 -or $null -eq $packages[0].PSObject.Properties['version']) {
    throw "MCP Registry metadata must contain exactly one versioned '$packageId' package."
}

$server.version = $Version
$packages[0].version = $Version
$content = ($server | ConvertTo-Json -Depth 20) -replace "`r?`n", "`n"
[System.IO.File]::WriteAllText(
    $ServerJsonPath,
    "$content`n",
    [System.Text.UTF8Encoding]::new($false))

$updated = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
$updatedPackages = @($updated.packages | Where-Object { $_.identifier -eq $packageId })
if ($updated.version -ne $Version -or
    $updatedPackages.Count -ne 1 -or
    $updatedPackages[0].version -ne $Version) {
    throw "MCP Registry metadata validation failed after stamping version '$Version'."
}

Write-Output "Updated MCP Registry metadata to version $Version."
