[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [switch]$SkillTests,
    [string[]]$ChangedPaths = @()
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$selections = [ordered]@{}

if ($Local) {
    if ($HookTests -or $SkillTests) {
        $selections['SkillGeneration'] = 'RequiresPowerPoint!=true'
    }
    if ($Contracts) {
        $selections['CLI'] = 'RequiresPowerPoint!=true'
        $selections['McpServer'] = 'RequiresPowerPoint!=true'
    }
    foreach ($path in $ChangedPaths) {
        if ($path -match '^tests[/\\]PowerPointMcp\.(CLI|McpServer|SkillGeneration)\.Tests[/\\]') {
            $selections[$Matches[1]] = 'RequiresPowerPoint!=true'
        }
        if ($path -match '^(vscode-extension|mcpb|npm-packages)[/\\]|^scripts[/\\](Build-VscodeExtension|Build-NpmPackages|Test-NpmPackages|PackageHelpers)\.ps1$|^\.github[/\\]workflows[/\\](ci|release)\.yml$') {
            $selections['SkillGeneration'] = 'RequiresPowerPoint!=true'
        }
        if ($path -match '^(\.github[/\\]plugins|plugins)[/\\]|^scripts[/\\](Build-Plugins|Sync-PublishedPluginRepo)\.ps1$|^\.github[/\\]workflows[/\\]publish-plugins\.yml$') {
            $selections['SkillGeneration'] = 'RequiresPowerPoint!=true'
        }
    }
}
else {
    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
        $selections[$project] = 'RequiresPowerPoint!=true'
    }
}

if ($selections.Count -eq 0) {
    throw 'No PowerPoint-free test group was selected.'
}

$results = Join-Path $root "TestResults\powerpoint-free-$([Guid]::NewGuid().ToString('N'))"
foreach ($entry in $selections.GetEnumerator()) {
    $project = Join-Path $root "tests\PowerPointMcp.$($entry.Key).Tests\PowerPointMcp.$($entry.Key).Tests.csproj"
    $info = [Diagnostics.ProcessStartInfo]::new('dotnet')
    $info.WorkingDirectory = $root
    $info.UseShellExecute = $false
    foreach ($argument in @(
        'test', $project, '-c', 'Release', '--no-build', '--no-restore',
        '--filter', $entry.Value,
        '--blame-hang-timeout', '5m',
        '--results-directory', $results,
        '--logger', "trx;LogFileName=$($entry.Key).trx",
        '-p:PowerPointMcpSkipCleanup=true'
    )) {
        $info.ArgumentList.Add($argument)
    }

    $process = [Diagnostics.Process]::Start($info)
    try {
        if (-not $process.WaitForExit(1800000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "$($entry.Key) tests exceeded the 30-minute deadline."
        }
        if ($process.ExitCode -ne 0) {
            throw "$($entry.Key) tests failed with exit code $($process.ExitCode)."
        }
    }
    finally {
        $process.Dispose()
    }

    $report = Join-Path $results "$($entry.Key).trx"
    if (-not (Test-Path -LiteralPath $report)) {
        throw "No test report for $($entry.Key)."
    }
    [xml]$trx = Get-Content -LiteralPath $report -Raw
    $counters = $trx.TestRun.ResultSummary.Counters
    if ([int]$counters.total -le 0 -or [int]$counters.passed -ne [int]$counters.total) {
        throw "$($entry.Key) selection was empty, skipped, or failed. See $report."
    }
}

$global:LASTEXITCODE = 0
