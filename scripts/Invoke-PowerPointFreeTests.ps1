[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [switch]$SkillTests,
    [string[]]$ChangedPaths = @(),
    [switch]$SelectChangedPaths,
    [switch]$ListOnly
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$selections = [ordered]@{}

if ($SelectChangedPaths) {
    if ($Local) {
        throw '-SelectChangedPaths cannot be combined with -Local.'
    }

    if ($ChangedPaths.Count -eq 0) {
        foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
            $selections[$project] = 'RequiresPowerPoint!=true'
        }
    }
    else {
        foreach ($path in $ChangedPaths) {
            $normalizedPath = $path.Replace('\', '/')
            switch -Regex ($normalizedPath) {
                '(^|/)[^/]+\.(csproj|props|targets)$' {
                    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
                        $selections[$project] = 'RequiresPowerPoint!=true'
                    }
                    break
                }
                '^tests/PowerPointMcp\.(CLI|McpServer|SkillGeneration)\.Tests/' {
                    $selections[$Matches[1]] = 'RequiresPowerPoint!=true'
                    break
                }
                '^tests/' {
                    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
                        $selections[$project] = 'RequiresPowerPoint!=true'
                    }
                    break
                }
                '^src/PowerPointMcp\.McpServer/|^src/PowerPointMcp\.Generators\.Mcp/' {
                    $selections['McpServer'] = 'RequiresPowerPoint!=true'
                    break
                }
                '^src/PowerPointMcp\.CLI/|^src/PowerPointMcp\.Generators\.Cli/' {
                    $selections['CLI'] = 'RequiresPowerPoint!=true'
                    break
                }
                '^src/PowerPointMcp\.(Core|ComInterop|Service|Generators\.Shared|Generators)(/|$)|^src/PowerPointMcp\.SkillGeneration/' {
                    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
                        $selections[$project] = 'RequiresPowerPoint!=true'
                    }
                    break
                }
                '^skills/|^vscode-extension/|^mcpb/|^npm-packages/' {
                    $selections['SkillGeneration'] = 'RequiresPowerPoint!=true'
                    break
                }
                '^scripts/|^\.github/workflows/|^(Sbroenne\.PowerPointMcp\.slnx|global\.json|Directory\..*|Directory\.Packages\.props)$' {
                    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
                        $selections[$project] = 'RequiresPowerPoint!=true'
                    }
                    break
                }
                '^(docs/|gh-pages/|\.changeset/|\.github/instructions/)|(^|/)[^/]+\.md$|^doc-counts\.json$' {
                    break
                }
                default {
                    foreach ($project in @('CLI', 'McpServer', 'SkillGeneration')) {
                        $selections[$project] = 'RequiresPowerPoint!=true'
                    }
                    break
                }
            }
        }
    }
}
elseif ($Local) {
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
    if ($SelectChangedPaths) {
        if ($ListOnly) {
            Write-Output 'No PowerPoint-free test groups selected.'
        }
        else {
            Write-Host 'No PowerPoint-free test groups are affected by the changed paths.'
        }
        $global:LASTEXITCODE = 0
        return
    }
    throw 'No PowerPoint-free test group was selected.'
}

if ($ListOnly) {
    $selections.Keys
    $global:LASTEXITCODE = 0
    return
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
