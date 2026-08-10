<#
.SYNOPSIS
    Records an exact daily star-count snapshot and generates the star-history SVG.

.DESCRIPTION
    Reads aggregate-only star history, replaces or appends the supplied UTC daily
    snapshot, persists the normalized aggregates, and renders a theme-aware SVG.
    Network access is intentionally handled by the caller.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$AggregatePath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\d{4}-\d{2}-\d{2}$")]
    [string]$SnapshotDate,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$SnapshotCount,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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

function ConvertTo-SnapshotDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $parsed = [DateTime]::MinValue
    if (-not [DateTime]::TryParseExact(
        $Value,
        "yyyy-MM-dd",
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal,
        [ref]$parsed)) {
        throw "Invalid star-history date '$Value'. Expected yyyy-MM-dd."
    }

    return $parsed
}

$resolvedAggregatePath = [System.IO.Path]::GetFullPath($AggregatePath)
if (-not (Test-Path -LiteralPath $resolvedAggregatePath -PathType Leaf)) {
    throw "Star-history aggregate file not found: $resolvedAggregatePath"
}

$rawHistory = Get-Content -LiteralPath $resolvedAggregatePath -Raw | ConvertFrom-Json
$snapshotsByDate = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::Ordinal)

foreach ($rawSnapshot in @($rawHistory)) {
    $propertyNames = @($rawSnapshot.PSObject.Properties.Name | Sort-Object)
    if (($propertyNames -join ",") -ne "count,date") {
        throw "Star-history records may contain only 'date' and 'count'."
    }

    $date = [string]$rawSnapshot.date
    [void](ConvertTo-SnapshotDate -Value $date)

    $countText = [string]$rawSnapshot.count
    $count = 0
    if ($countText -notmatch "^\d+$" -or -not [int]::TryParse(
        $countText,
        [System.Globalization.NumberStyles]::None,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$count)) {
        throw "Invalid star count '$countText' for $date."
    }

    if (-not $snapshotsByDate.TryAdd($date, $count)) {
        throw "Duplicate star-history date '$date'."
    }
}

[void](ConvertTo-SnapshotDate -Value $SnapshotDate)
$snapshotsByDate[$SnapshotDate] = $SnapshotCount

$snapshots = @(
    $snapshotsByDate.GetEnumerator() |
        Sort-Object -Property Key |
        ForEach-Object {
            [ordered]@{
                date = $_.Key
                count = $_.Value
            }
        }
)

if ($snapshots.Count -eq 0) {
    throw "Star history must contain at least one snapshot."
}

$aggregateDirectory = Split-Path -Parent $resolvedAggregatePath
if (-not (Test-Path -LiteralPath $aggregateDirectory)) {
    New-Item -ItemType Directory -Path $aggregateDirectory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$aggregateJson = $snapshots | ConvertTo-Json -Depth 3 -AsArray
[System.IO.File]::WriteAllText($resolvedAggregatePath, "$aggregateJson`n", $utf8NoBom)

$firstDate = ConvertTo-SnapshotDate -Value $snapshots[0].date
$lastDate = ConvertTo-SnapshotDate -Value $snapshots[-1].date
$chartEnd = if ($lastDate -eq $firstDate) { $firstDate.AddDays(1) } else { $lastDate }

$width = 900
$height = 480
$left = 72
$right = 24
$top = 76
$bottom = 62
$plotWidth = $width - $left - $right
$plotHeight = $height - $top - $bottom
$durationTicks = ($chartEnd - $firstDate).Ticks
$maxStars = [Math]::Max(1, [int](($snapshots | Measure-Object -Property count -Maximum).Maximum))

$points = @(
    foreach ($snapshot in $snapshots) {
        $date = ConvertTo-SnapshotDate -Value $snapshot.date
        $x = $left + ((($date - $firstDate).Ticks / $durationTicks) * $plotWidth)
        $y = $top + $plotHeight - (($snapshot.count / $maxStars) * $plotHeight)

        [pscustomobject]@{
            X = $x
            Y = $y
        }
    }
)

$linePath = "M $(ConvertTo-SvgNumber $points[0].X) $(ConvertTo-SvgNumber $points[0].Y)"
for ($index = 1; $index -lt $points.Count; $index++) {
    $linePath += " H $(ConvertTo-SvgNumber $points[$index].X) V $(ConvertTo-SvgNumber $points[$index].Y)"
}

$firstX = ConvertTo-SvgNumber $points[0].X
$lastX = ConvertTo-SvgNumber $points[-1].X
$baselineY = ConvertTo-SvgNumber ($top + $plotHeight)
$areaPath = "M $firstX $baselineY L $firstX $(ConvertTo-SvgNumber $points[0].Y)"
for ($index = 1; $index -lt $points.Count; $index++) {
    $areaPath += " H $(ConvertTo-SvgNumber $points[$index].X) V $(ConvertTo-SvgNumber $points[$index].Y)"
}
$areaPath += " L $lastX $baselineY Z"

$repositoryText = ConvertTo-SvgText $Repository
$dateRange = "{0:MMM yyyy} - {1:MMM yyyy}" -f $firstDate, $lastDate
$currentCount = [int]$snapshots[-1].count
$subtitle = ConvertTo-SvgText "$Repository - $currentCount stars - $dateRange"
$description = ConvertTo-SvgText (
    "Exact cumulative GitHub star snapshots for $Repository from " +
    "$($firstDate.ToString('yyyy-MM-dd')) to $($lastDate.ToString('yyyy-MM-dd')).")

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
    $tickDate = $firstDate.AddTicks([long](($durationTicks * $index) / 4))
    $anchor = if ($index -eq 0) { "start" } elseif ($index -eq 4) { "end" } else { "middle" }

    [void]$svg.AppendLine("  <text class=`"axis-text`" x=`"$(ConvertTo-SvgNumber $x)`" y=`"$($top + $plotHeight + 28)`" text-anchor=`"$anchor`">$($tickDate.ToString('MMM yyyy'))</text>")
}

[void]$svg.AppendLine("  <path class=`"area`" d=`"$areaPath`" />")
[void]$svg.AppendLine("  <path class=`"line`" d=`"$linePath`" />")
[void]$svg.AppendLine("</svg>")

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

[System.IO.File]::WriteAllText($resolvedOutputPath, $svg.ToString(), $utf8NoBom)
Write-Host "Recorded $SnapshotDate at $SnapshotCount stars and generated $resolvedOutputPath." -ForegroundColor Green
