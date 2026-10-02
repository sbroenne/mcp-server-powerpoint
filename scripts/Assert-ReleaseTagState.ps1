[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$')]
    [string] $Version,

    [string] $RepositoryRoot = (Split-Path -Parent $PSScriptRoot),

    [switch] $ResumeRelease,

    [switch] $CreateTag
)

$ErrorActionPreference = 'Stop'
$tag = "v$Version"

& git -C $RepositoryRoot rev-parse --verify "refs/tags/$tag" 2>$null | Out-Null
$tagExists = $LASTEXITCODE -eq 0

if ($tagExists) {
    if (-not $ResumeRelease) {
        throw "Tag $tag already exists. Set resume_release to resume a partial release."
    }

    Write-Output "Resume requested, and tag $tag already exists."
    exit 0
}

if ($ResumeRelease) {
    throw "Cannot resume because tag $tag does not exist."
}

if (-not $CreateTag) {
    Write-Output "Tag $tag is available for a new release."
    exit 0
}

Write-Output "Creating tag: $tag"
& git -C $RepositoryRoot tag -a $tag -m "Release $tag"
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create tag $tag."
}

& git -C $RepositoryRoot push origin $tag
if ($LASTEXITCODE -ne 0) {
    throw "Failed to push tag $tag."
}
