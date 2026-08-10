$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Assert-Equal {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [object]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Condition,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repositoryRoot "scripts/Update-StarHistory.ps1"
$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) "powerpointmcp-star-history-$([Guid]::NewGuid().ToString('N'))"
$aggregatePath = Join-Path $testDirectory "star-history.json"
$svgPath = Join-Path $testDirectory "star-history.svg"

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null

    @(
        [ordered]@{ date = "2026-07-01"; count = 1 }
        [ordered]@{ date = "2026-07-02"; count = 3 }
    ) | ConvertTo-Json -AsArray | Set-Content -Path $aggregatePath -Encoding utf8NoBOM

    & $scriptPath `
        -Repository "sbroenne/mcp-server-powerpoint" `
        -AggregatePath $aggregatePath `
        -SnapshotDate "2026-07-03" `
        -SnapshotCount 2 `
        -OutputPath $svgPath

    $history = @(Get-Content -Raw -Path $aggregatePath | ConvertFrom-Json)
    Assert-Equal -Actual $history.Count -Expected 3 -Message "A new UTC day should append one snapshot."
    Assert-Equal -Actual $history[-1].date -Expected "2026-07-03" -Message "Snapshots should be date ordered."
    Assert-Equal -Actual $history[-1].count -Expected 2 -Message "Star counts may decrease after an unstar."

    foreach ($snapshot in $history) {
        $propertyNames = @($snapshot.PSObject.Properties.Name | Sort-Object)
        Assert-Equal -Actual ($propertyNames -join ",") -Expected "count,date" -Message "Persistent snapshots must contain aggregate fields only."
    }

    $svg = Get-Content -Raw -Path $svgPath
    Assert-True -Condition ($svg.Contains("2 stars")) -Message "The SVG should label the exact current count."
    Assert-True -Condition (-not $svg.Contains("approx")) -Message "The SVG must not describe the data as approximate."
    Assert-True -Condition (-not $svg.Contains("WatchEvent")) -Message "The SVG must not reference WatchEvent data."

    & $scriptPath `
        -Repository "sbroenne/mcp-server-powerpoint" `
        -AggregatePath $aggregatePath `
        -SnapshotDate "2026-07-03" `
        -SnapshotCount 4 `
        -OutputPath $svgPath

    $history = @(Get-Content -Raw -Path $aggregatePath | ConvertFrom-Json)
    Assert-Equal -Actual $history.Count -Expected 3 -Message "A rerun on the same UTC day must replace, not duplicate, its snapshot."
    Assert-Equal -Actual $history[-1].count -Expected 4 -Message "A same-day rerun should use the latest exact count."

    $rejectedNegativeCount = $false
    try {
        & $scriptPath `
            -Repository "sbroenne/mcp-server-powerpoint" `
            -AggregatePath $aggregatePath `
            -SnapshotDate "2026-07-04" `
            -SnapshotCount -1 `
            -OutputPath $svgPath
    }
    catch {
        $rejectedNegativeCount = $true
    }

    Assert-True -Condition $rejectedNegativeCount -Message "Negative star counts must be rejected."
    Write-Host "Star history tests passed." -ForegroundColor Green
}
finally {
    if (Test-Path $testDirectory) {
        Remove-Item -Path $testDirectory -Recurse -Force
    }
}
