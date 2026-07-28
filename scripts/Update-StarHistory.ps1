<#
.SYNOPSIS
    Generates an SVG chart from a repository's GitHub watch-event history.

.DESCRIPTION
    Reads repository metadata and public watch events from GitHub's API and writes a
    deterministic, theme-aware SVG suitable for the repository README and docs site.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$Token,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function ConvertTo-SvgText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-SvgNumber {
    param(
        [Parameter(Mandatory = $true)]
        [double]$Value
    )

    return $Value.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-DateTimeOffset {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [DateTimeOffset]) {
        return $Value
    }

    if ($Value -is [DateTime]) {
        return [DateTimeOffset]$Value
    }

    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            $converted = ConvertTo-DateTimeOffset -Value $item
            if ($null -ne $converted) {
                return $converted
            }
        }

        return $null
    }

    if ($Value -is [string]) {
        $trimmed = $Value.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            return $null
        }

        return [DateTimeOffset]::Parse($trimmed, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
    }

    return [DateTimeOffset]$Value
}

function Invoke-GitHubApi {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter()]
        [bool]$AllowFallback = $false,

        [Parameter()]
        [bool]$IsAuthenticated = $false
    )

    try {
        return Invoke-RestMethod -Uri $Uri -Headers $Headers
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if ($AllowFallback -and $IsAuthenticated -and ($statusCode -eq 401 -or $statusCode -eq 403)) {
            Write-Warning "GitHub rejected the supplied token for '$Uri'; retrying without authentication."

            $unauthHeaders = @{}
            foreach ($key in $Headers.Keys) {
                if ($key -ne 'Authorization') {
                    $unauthHeaders[$key] = $Headers[$key]
                }
            }

            return Invoke-RestMethod -Uri $Uri -Headers $unauthHeaders
        }

        throw
    }
}

$headers = @{
    Accept = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent" = "sbroenne/mcp-server-powerpoint-star-history"
}

$isAuthenticatedRequest = -not [string]::IsNullOrWhiteSpace($Token)
if ($isAuthenticatedRequest) {
    $headers.Authorization = "Bearer $Token"
}

$eventsUri = "https://api.github.com/repos/$Repository/events?per_page=100"
$events = @()
$repoMetadata = $null

try {
    $repoMetadata = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Repository" -Headers $headers -AllowFallback $true -IsAuthenticated $isAuthenticatedRequest
    $rawEvents = @(Invoke-GitHubApi -Uri $eventsUri -Headers $headers -AllowFallback $true -IsAuthenticated $isAuthenticatedRequest)
    $events = @($rawEvents | Where-Object { $_.type -eq "WatchEvent" })
}
catch {
    Write-Warning "Unable to load repository watch events for '$Repository'. $($_.Exception.Message)"
}

$stargazers = [System.Collections.Generic.List[object]]::new()

if ($repoMetadata) {
    $createdAt = ConvertTo-DateTimeOffset -Value $repoMetadata.created_at
    if ($null -ne $createdAt) {
        $stargazers.Add($createdAt)
    }
}

foreach ($event in $events) {
    $createdAt = ConvertTo-DateTimeOffset -Value $event.created_at
    if ($null -ne $createdAt) {
        $stargazers.Add($createdAt)
    }
}

$stars = @($stargazers | Where-Object { $null -ne $_ } | Sort-Object)

if ($stars.Count -eq 0) {
    throw "No star-related events were returned for '$Repository'."
}
$firstStar = $stars[0]
$lastStar = $stars[-1]
$chartEnd = $lastStar

if ($chartEnd -eq $firstStar) {
    $chartEnd = $firstStar.AddDays(1)
}

$width = 900
$height = 480
$left = 72
$right = 24
$top = 76
$bottom = 62
$plotWidth = $width - $left - $right
$plotHeight = $height - $top - $bottom
$durationTicks = ($chartEnd - $firstStar).Ticks
$maxStars = $stars.Count

if ($repoMetadata -and $repoMetadata.stargazers_count -gt $maxStars) {
    $maxStars = [int]$repoMetadata.stargazers_count
}

if ($maxStars -lt 1) {
    $maxStars = 1
}

$points = for ($index = 0; $index -lt $stars.Count; $index++) {
    $elapsedTicks = ($stars[$index] - $firstStar).Ticks
    $x = $left + (($elapsedTicks / $durationTicks) * $plotWidth)
    $y = $top + $plotHeight - ((($index + 1) / $maxStars) * $plotHeight)

    [pscustomobject]@{
        X = $x
        Y = $y
    }
}

$lineCoordinates = ($points | ForEach-Object {
    "$(ConvertTo-SvgNumber $_.X) $(ConvertTo-SvgNumber $_.Y)"
}) -join " L "
$linePath = "M $lineCoordinates"

$firstX = ConvertTo-SvgNumber $points[0].X
$lastX = ConvertTo-SvgNumber $points[-1].X
$baselineY = ConvertTo-SvgNumber ($top + $plotHeight)
$areaPath = "M $firstX $baselineY L $lineCoordinates L $lastX $baselineY Z"

$repositoryText = ConvertTo-SvgText $Repository
$dateRange = "{0:MMM yyyy} - {1:MMM yyyy}" -f $firstStar, $lastStar
$subtitle = ConvertTo-SvgText "$Repository - watch events / approx. $maxStars stars - $dateRange"
$description = ConvertTo-SvgText (
    "Cumulative GitHub stars for $Repository from " +
    "$($firstStar.ToString('yyyy-MM-dd')) to $($lastStar.ToString('yyyy-MM-dd')).")

$svg = [System.Text.StringBuilder]::new()
[void]$svg.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$svg.AppendLine("<svg xmlns=`"http://www.w3.org/2000/svg`" width=`"$width`" height=`"$height`" viewBox=`"0 0 $width $height`" role=`"img`" aria-labelledby=`"title description`">")
[void]$svg.AppendLine("  <title id=`"title`">GitHub stars over time for $repositoryText</title>")
[void]$svg.AppendLine("  <desc id=`"description`">$description</desc>")
[void]$svg.AppendLine("  <style>")
[void]$svg.AppendLine("    .background { fill: #ffffff; }")
[void]$svg.AppendLine("    .grid { stroke: #d0d7de; stroke-width: 1; }")
[void]$svg.AppendLine("    .axis-text { fill: #57606a; font: 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .title { fill: #1f2328; font: 600 22px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .subtitle { fill: #57606a; font: 14px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .area { fill: #2da44e; opacity: 0.14; }")
[void]$svg.AppendLine("    .line { fill: none; stroke: #1f883d; stroke-linecap: round; stroke-linejoin: round; stroke-width: 3; }")
[void]$svg.AppendLine("    @media (prefers-color-scheme: dark) {")
[void]$svg.AppendLine("      .background { fill: #0d1117; }")
[void]$svg.AppendLine("      .grid { stroke: #30363d; }")
[void]$svg.AppendLine("      .axis-text, .subtitle { fill: #8b949e; }")
[void]$svg.AppendLine("      .title { fill: #f0f6fc; }")
[void]$svg.AppendLine("      .area { fill: #3fb950; opacity: 0.18; }")
[void]$svg.AppendLine("      .line { stroke: #3fb950; }")
[void]$svg.AppendLine("    }")
[void]$svg.AppendLine("  </style>")
[void]$svg.AppendLine("  <rect class=`"background`" width=`"$width`" height=`"$height`" rx=`"8`" />")
[void]$svg.AppendLine("  <text class=`"title`" x=`"$left`" y=`"34`">GitHub stars over time</text>")
[void]$svg.AppendLine("  <text class=`"subtitle`" x=`"$left`" y=`"57`">$subtitle</text>")

for ($index = 0; $index -le 4; $index++) {
    $value = [Math]::Round(($maxStars * $index) / 4)
    $y = $top + $plotHeight - (($value / $maxStars) * $plotHeight)
    $yText = ConvertTo-SvgNumber $y

    [void]$svg.AppendLine("  <line class=`"grid`" x1=`"$left`" y1=`"$yText`" x2=`"$($left + $plotWidth)`" y2=`"$yText`" />")
    [void]$svg.AppendLine("  <text class=`"axis-text`" x=`"$($left - 12)`" y=`"$yText`" text-anchor=`"end`" dominant-baseline=`"middle`">$value</text>")
}

for ($index = 0; $index -le 4; $index++) {
    $x = $left + (($plotWidth * $index) / 4)
    $tickDate = $firstStar.AddTicks([long](($durationTicks * $index) / 4))
    $anchor = if ($index -eq 0) { "start" } elseif ($index -eq 4) { "end" } else { "middle" }

    [void]$svg.AppendLine("  <text class=`"axis-text`" x=`"$(ConvertTo-SvgNumber $x)`" y=`"$($top + $plotHeight + 28)`" text-anchor=`"$anchor`">$($tickDate.ToString('MMM yyyy'))</text>")
}

[void]$svg.AppendLine("  <path class=`"area`" d=`"$areaPath`" />")
[void]$svg.AppendLine("  <path class=`"line`" d=`"$linePath`" />")
[void]$svg.AppendLine("</svg>")

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath

if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($resolvedOutputPath, $svg.ToString(), $utf8NoBom)

Write-Host "Generated $resolvedOutputPath with $maxStars stars." -ForegroundColor Green
