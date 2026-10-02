<#
.SYNOPSIS
    Builds and verifies native Windows x64 and ARM64 VS Code extension packages.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = '.\artifacts\vscode',
    [string]$SkillsDirectory,
    [string]$X64Runtime,
    [string]$Arm64Runtime
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

function Invoke-Checked {
    param([string]$Name, [scriptblock]$Action)
    Write-Host $Name -ForegroundColor Cyan
    $global:LASTEXITCODE = 0
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed with exit code $LASTEXITCODE." }
}

function Read-ZipEntry {
    param([IO.Compression.ZipArchive]$Archive, [string]$Name)
    $entry = $Archive.GetEntry($Name)
    if (-not $entry) { throw "VSIX is missing $Name." }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

if (-not $Version) {
    [xml]$props = Get-Content (Join-Path $root 'Directory.Build.props')
    $Version = $props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
}
if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$') {
    throw 'A valid package version is required.'
}

$output = [IO.Path]::GetFullPath($OutputDir, $root)
Assert-PackageOutputPath -Path $output -RepoRoot $root -Inputs @($SkillsDirectory, $X64Runtime, $Arm64Runtime)
if (Test-Path -LiteralPath $output) {
    throw "Use a new VS Code package output directory: $output"
}
New-Item -ItemType Directory -Path $output | Out-Null

$stage = Join-Path ([IO.Path]::GetTempPath()) "PowerPointMcpExtension-$([Guid]::NewGuid().ToString('N'))"
$runtimeRoot = Join-Path $output 'runtimes'
try {
    if (-not $X64Runtime) {
        $runtimeDir = Join-Path $runtimeRoot 'x64'
        Publish-PackageRuntime -Component Mcp -RepoRoot $root -Version $Version `
            -Architecture x64 -OutputDirectory $runtimeDir
        $X64Runtime = Join-Path $runtimeDir 'Sbroenne.PowerPointMcp.McpServer.exe'
    }
    if (-not $Arm64Runtime) {
        $runtimeDir = Join-Path $runtimeRoot 'arm64'
        Publish-PackageRuntime -Component Mcp -RepoRoot $root -Version $Version `
            -Architecture arm64 -OutputDirectory $runtimeDir
        $Arm64Runtime = Join-Path $runtimeDir 'Sbroenne.PowerPointMcp.McpServer.exe'
    }
    $X64Runtime = (Resolve-Path -LiteralPath $X64Runtime).Path
    $Arm64Runtime = (Resolve-Path -LiteralPath $Arm64Runtime).Path
    Assert-PackageRuntimeArchitecture -Path $X64Runtime -Architecture x64
    Assert-PackageRuntimeArchitecture -Path $Arm64Runtime -Architecture arm64

    if (-not $SkillsDirectory) {
        $SkillsDirectory = Join-Path $root 'skills'
    }

    New-Item -ItemType Directory -Path $stage | Out-Null
    Get-ChildItem (Join-Path $root 'vscode-extension') -Force |
        Where-Object {
            $_.Name -notin @('node_modules', 'bin', 'out', 'skills') -and
            $_.Extension -ne '.vsix'
        } |
        Copy-Item -Destination $stage -Recurse
    $bin = New-Item -ItemType Directory -Path (Join-Path $stage 'bin')
    $skills = New-Item -ItemType Directory -Path (Join-Path $stage 'skills')
    Copy-Item (Join-Path $SkillsDirectory 'powerpoint-mcp') $skills.FullName -Recurse
    Set-Content (Join-Path $skills.FullName 'powerpoint-mcp\VERSION') $Version -NoNewline
    Copy-Item (Join-Path $root 'CHANGELOG.md') $stage -Force
    Copy-Item (Join-Path $root 'LICENSE') $stage -Force

    $manifestPath = Join-Path $stage 'package.json'
    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $manifest.version = $Version
    $manifest.scripts.'vscode:prepublish' = 'npm run compile'
    $manifest | ConvertTo-Json -Depth 20 | Set-Content $manifestPath -Encoding utf8

    $packages = @(
        @{
            Target = 'win32-x64'
            Architecture = 'x64'
            Runtime = $X64Runtime
            FileName = "powerpoint-mcp-$Version.vsix"
        },
        @{
            Target = 'win32-arm64'
            Architecture = 'arm64'
            Runtime = $Arm64Runtime
            FileName = "powerpoint-mcp-$Version-win32-arm64.vsix"
        }
    )

    Push-Location $stage
    try {
        Invoke-Checked 'Install extension dependencies' { npm.cmd ci --ignore-scripts }
        Invoke-Checked 'Compile extension' { npm.cmd run compile }
        Invoke-Checked 'Lint extension' { npm.cmd run lint }
        Invoke-Checked 'Type-check extension tests' { npm.cmd run typecheck:tests }
        Invoke-Checked 'Test extension' { npm.cmd test }

        foreach ($package in $packages) {
            Copy-Item -LiteralPath $package.Runtime -Destination $bin.FullName -Force
            Invoke-Checked "Package extension ($($package.Target))" {
                npm.cmd exec -- vsce package --no-dependencies --target $package.Target `
                    --out (Join-Path $output $package.FileName)
            }
        }
    }
    finally {
        Pop-Location
    }

    foreach ($package in $packages) {
        $packagePath = Join-Path $output $package.FileName
        $vsix = [IO.Compression.ZipFile]::OpenRead($packagePath)
        try {
            foreach ($required in @(
                'extension/bin/Sbroenne.PowerPointMcp.McpServer.exe',
                'extension/out/extension.js',
                'extension/out/prerequisites.js',
                'extension/skills/powerpoint-mcp/SKILL.md',
                'extension/skills/powerpoint-mcp/VERSION'
            )) {
                if (-not $vsix.GetEntry($required)) { throw "VSIX is missing $required." }
            }
            $skillSource = Join-Path $SkillsDirectory 'powerpoint-mcp'
            foreach ($skillFile in Get-ChildItem -LiteralPath $skillSource -File -Recurse) {
                $relative = [IO.Path]::GetRelativePath($skillSource, $skillFile.FullName).Replace('\', '/')
                if (-not $vsix.GetEntry("extension/skills/powerpoint-mcp/$relative")) {
                    throw "VSIX is missing skill file $relative."
                }
            }

            $inspectionRuntime = Join-Path $output "$($package.Target)-server-inspection.exe"
            try {
                [IO.Compression.ZipFileExtensions]::ExtractToFile(
                    $vsix.GetEntry('extension/bin/Sbroenne.PowerPointMcp.McpServer.exe'),
                    $inspectionRuntime,
                    $false)
                Assert-PackageRuntimeArchitecture -Path $inspectionRuntime -Architecture $package.Architecture
            }
            finally {
                if (Test-Path -LiteralPath $inspectionRuntime) {
                    Remove-Item -LiteralPath $inspectionRuntime -Force
                }
            }

            if ((Read-ZipEntry $vsix 'extension/skills/powerpoint-mcp/VERSION').Trim() -ne $Version) {
                throw 'VSIX skill version does not match the package.'
            }
            $packagedManifest = Read-ZipEntry $vsix 'extension/package.json' | ConvertFrom-Json
            if ($packagedManifest.version -ne $Version -or
                ($packagedManifest.extensionKind -join ',') -ne 'ui' -or
                ($packagedManifest.os -join ',') -ne 'win32') {
                throw 'VSIX version or local Windows host metadata is incorrect.'
            }
            [xml]$metadata = Read-ZipEntry $vsix 'extension.vsixmanifest'
            if ($metadata.PackageManifest.Metadata.Identity.TargetPlatform -ne $package.Target) {
                throw "VSIX target does not match $($package.Target)."
            }
            foreach ($entry in $vsix.Entries) {
                if ($entry.FullName -match '^extension/(node_modules|tests|scripts|coverage|out/tests)/' -or
                    $entry.FullName -match '^extension/(vitest\.config\.|tsconfig(?:\.test)?\.json$)') {
                    throw "VSIX contains development files: $($entry.FullName)"
                }
            }
        }
        finally {
            $vsix.Dispose()
        }
        Write-Host "Verified VSIX: $($package.FileName) ($($package.Target), $($package.Architecture))"
    }
}
finally {
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
}
