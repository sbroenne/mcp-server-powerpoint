<#
.SYNOPSIS
    Bootstraps exact aggregate star history with the authenticated GitHub CLI.

.DESCRIPTION
    Pages through repository.stargazers edges.starredAt via `gh api graphql`.
    The output contains only UTC date/count aggregates; account names, node IDs,
    and pagination cursors are never persisted.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is required."
}

$owner, $name = $Repository.Split("/", 2)
$query = @'
query($owner: String!, $name: String!, $cursor: String) {
  repository(owner: $owner, name: $name) {
    createdAt
    stargazerCount
    stargazers(first: 100, after: $cursor, orderBy: { field: STARRED_AT, direction: ASC }) {
      pageInfo {
        hasNextPage
        endCursor
      }
      edges {
        starredAt
      }
    }
  }
}
'@

$starredAt = [System.Collections.Generic.List[DateTimeOffset]]::new()
$cursor = $null
$createdAt = $null
$expectedCount = $null

do {
    $arguments = @(
        "api"
        "graphql"
        "-f", "query=$query"
        "-f", "owner=$owner"
        "-f", "name=$name"
    )

    if ($null -ne $cursor) {
        $arguments += @("-f", "cursor=$cursor")
    }

    $responseText = & gh @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "gh api graphql failed while collecting star history for '$Repository'."
    }

    $response = $responseText | ConvertFrom-Json
    $errorsProperty = $response.PSObject.Properties["errors"]
    if ($null -ne $errorsProperty -and $null -ne $errorsProperty.Value) {
        $messages = @($errorsProperty.Value | ForEach-Object { $_.message }) -join "; "
        throw "GitHub GraphQL returned errors: $messages"
    }

    $repositoryData = $response.data.repository
    if ($null -eq $repositoryData) {
        throw "Repository '$Repository' was not returned by GitHub GraphQL."
    }

    if ($null -eq $createdAt) {
        $createdAt = [DateTimeOffset]::Parse(
            [string]$repositoryData.createdAt,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        $expectedCount = [int]$repositoryData.stargazerCount
    }

    foreach ($edge in @($repositoryData.stargazers.edges)) {
        $starredAt.Add([DateTimeOffset]::Parse(
            [string]$edge.starredAt,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal))
    }

    $pageInfo = $repositoryData.stargazers.pageInfo
    $cursor = if ($pageInfo.hasNextPage) { [string]$pageInfo.endCursor } else { $null }
} while ($null -ne $cursor)

if ($starredAt.Count -ne $expectedCount) {
    throw "GraphQL returned $($starredAt.Count) stargazer edges, but stargazerCount is $expectedCount."
}

$historyByDate = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::Ordinal)
if ($starredAt.Count -eq 0) {
    $historyByDate[$createdAt.UtcDateTime.ToString("yyyy-MM-dd")] = 0
}
else {
    $runningCount = 0
    foreach ($timestamp in @($starredAt | Sort-Object)) {
        $runningCount++
        $historyByDate[$timestamp.UtcDateTime.ToString("yyyy-MM-dd")] = $runningCount
    }
}

$history = @(
    $historyByDate.GetEnumerator() |
        Sort-Object -Property Key |
        ForEach-Object {
            [ordered]@{
                date = $_.Key
                count = $_.Value
            }
        }
)

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$json = $history | ConvertTo-Json -Depth 3 -AsArray
[System.IO.File]::WriteAllText($resolvedOutputPath, "$json`n", $utf8NoBom)

Write-Host "Wrote $($history.Count) aggregate records for $expectedCount exact stargazers to $resolvedOutputPath." -ForegroundColor Green
