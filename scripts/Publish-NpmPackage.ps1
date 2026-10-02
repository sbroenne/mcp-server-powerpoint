[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $PackageName,

    [Parameter(Mandatory)]
    [ValidatePattern('^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$')]
    [string] $Version,

    [Parameter(Mandatory)]
    [string] $PackageTarball,

    [string] $NpmCommand = 'npm'
)

$ErrorActionPreference = 'Stop'

$publishedVersion = & $NpmCommand view "$PackageName@$Version" version 2>$null
if ($LASTEXITCODE -eq 0 -and $publishedVersion -eq $Version) {
    Write-Output "Skipping $PackageName@$Version because it is already published."
    exit 0
}

& $NpmCommand publish $PackageTarball --access public
if ($LASTEXITCODE -ne 0) {
    throw "Failed to publish $PackageName@$Version."
}
