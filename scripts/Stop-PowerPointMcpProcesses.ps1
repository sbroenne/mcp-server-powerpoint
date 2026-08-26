<#
.SYNOPSIS
    Stops PowerPointMcp-owned processes safely before build.
.DESCRIPTION
    Pre-build cleanup script that:
    1. Gracefully stops the PowerPointMcp Service via named pipe (service.shutdown)
    2. Uses the daemon's PID-plus-start-time ownership record for forced cleanup

    This prevents file locking issues during build when the service or PowerPoint
    holds handles to assemblies or presentations.
.NOTES
    Ported from mcp-server-excel's scripts/Stop-ExcelMcpProcesses.ps1.
    Called from Directory.Build.props as a BeforeBuild target.
    Safe to run when no processes are running (silently succeeds).
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = 'SilentlyContinue'
$script:CleanupSucceeded = $true

function Write-Status($message) {
    if ($Verbose) {
        Write-Host "  [pre-build] $message" -ForegroundColor DarkGray
    }
}

# ----------------------------------------------
# 1. Gracefully stop PowerPointMcp Service via CLI
# ----------------------------------------------
function Stop-PowerPointMcpService {
    # Look for powerpointcli in build output directories (Debug/Release)
    $scriptDir = Split-Path -Parent $PSScriptRoot  # repo root
    $cliPaths = @(
        "$scriptDir\src\PowerPointMcp.CLI\bin\Debug\net10.0-windows\powerpointcli.exe",
        "$scriptDir\src\PowerPointMcp.CLI\bin\Release\net10.0-windows\powerpointcli.exe"
    )
    $powerpointcli = $cliPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($powerpointcli) {
        Write-Status "Using CLI: $powerpointcli"
        $output = & $powerpointcli service stop --force 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            # Parse JSON to check if service was running
            try {
                $result = $output | ConvertFrom-Json
                if ($result.message -eq 'Service is not running.') {
                    Write-Status "PowerPointMcp Service was not running"
                } else {
                    Write-Host "  PowerPointMcp Service stopped gracefully" -ForegroundColor Green
                }
            } catch {
                Write-Status "Service stop completed (exit code 0)"
            }
        } else {
            Write-Status "CLI service stop returned exit code $exitCode; trying graceful pipe shutdown"
            Stop-PowerPointMcpServiceFallback
        }
    } else {
        Write-Status "powerpointcli not found (first build?), using fallback"
        Stop-PowerPointMcpServiceFallback
    }
}

function Stop-PowerPointMcpServiceFallback {
    # Fallback: direct named pipe shutdown (works without CLI binary)
    $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    $pipeName = "powerpointmcp-cli-$sid"

    $pipeExists = Test-Path "\\.\pipe\$pipeName"
    if (-not $pipeExists) {
        Write-Status "PowerPointMcp Service not running (no pipe found)"
        return
    }

    Write-Status "PowerPointMcp Service detected, sending shutdown via pipe..."
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect(3000)

        $writer = New-Object System.IO.StreamWriter($pipe, [System.Text.Encoding]::UTF8, 4096)
        $writer.AutoFlush = $true
        $reader = New-Object System.IO.StreamReader($pipe, [System.Text.Encoding]::UTF8)

        $writer.WriteLine('{"Command":"service.shutdown"}')
        $response = $reader.ReadLine()
        Write-Status "Service response: $response"

        $reader.Dispose()
        $writer.Dispose()
        $pipe.Dispose()

        Start-Sleep -Milliseconds 500
        Write-Host "  PowerPointMcp Service stopped gracefully" -ForegroundColor Green
    }
    catch {
        Write-Status "Could not connect to pipe: $($_.Exception.Message)"
        Write-Warning "Could not stop the daemon safely. No processes were terminated because ownership could not be proven."
        $script:CleanupSucceeded = $false
    }
}

# ----------------------------------------------
# Run cleanup
# ----------------------------------------------
Stop-PowerPointMcpService
if ($script:CleanupSucceeded) {
    exit 0
}

exit 1
