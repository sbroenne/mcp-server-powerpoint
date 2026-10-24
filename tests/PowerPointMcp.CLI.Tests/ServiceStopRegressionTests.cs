using System.Diagnostics;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class ServiceStopRegressionTests
{
    [Theory]
    [InlineData("status")]
    [InlineData("stop")]
    public async Task ServiceControl_DoesNotDependOnPowerPointTrackingDiagnostics(string action)
    {
        string pipeName = $"PowerPointMcp_TrackingWarning_{Guid.NewGuid():N}";
        using var daemon = StartCli("service", "run", "--pipe-name", pipeName);
        try
        {
            await WaitForDaemonAsync(pipeName);
            string trackingPath = DaemonProcessTracker.GetTrackingFilePath(pipeName);
            File.WriteAllText(trackingPath, "{invalid-record");
            using var command = StartCli("service", action, "--pipe-name", pipeName);
            string output = await command.StandardOutput.ReadToEndAsync();
            await command.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.True(command.ExitCode == 0, output);
            Assert.Contains("powerPointProcessWarning", output, StringComparison.Ordinal);
            if (action == "status")
                Assert.Contains("\"responsive\":true", output, StringComparison.Ordinal);
        }
        finally
        {
            await StopTestServiceAsync(daemon, pipeName);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ForcedStop_StopsRecordedServiceEvenAfterItsLockWasReleased()
    {
        string pipeName = $"PowerPointMcp_ReleasedLock_{Guid.NewGuid():N}";
        using var child = Process.Start(new ProcessStartInfo("pwsh", "-NoProfile -Command \"Start-Sleep 60\"")
        {
            UseShellExecute = false,
            CreateNoWindow = true
        }) ?? throw new InvalidOperationException("Could not start the test service process.");
        try
        {
            DaemonProcessTracker.RegisterProcess(pipeName, child.Id,
                child.StartTime.ToUniversalTime().ToFileTimeUtc());
            using var stop = StartCli("service", "stop", "--force", "--pipe-name", pipeName);
            await stop.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.Equal(0, stop.ExitCode);
            var exit = child.WaitForExitAsync();
            await Task.WhenAny(exit, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.True(exit.IsCompleted, "An explicit forced stop must also stop a verified service that no longer holds its lock.");
        }
        finally
        {
            if (!child.HasExited)
            {
                child.Kill();
                await child.WaitForExitAsync();
            }
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ServiceStart_UsesExplicitPipeInsteadOfEnvironmentDefault()
    {
        string defaultPipe = $"PowerPointMcp_Default_{Guid.NewGuid():N}";
        string selectedPipe = $"PowerPointMcp_Selected_{Guid.NewGuid():N}";
        using var defaultDaemon = StartCli("service", "run", "--pipe-name", defaultPipe);
        using var selectedDaemon = StartCli("service", "run", "--pipe-name", selectedPipe);
        try
        {
            await WaitForDaemonAsync(defaultPipe);
            await WaitForDaemonAsync(selectedPipe);
            using var start = StartCliWithDefaultPipe(defaultPipe, "service", "start", "--pipe-name", selectedPipe);
            string output = await start.StandardOutput.ReadToEndAsync();
            await start.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(0, start.ExitCode);
            Assert.Contains($"\"processId\":{selectedDaemon.Id}", output, StringComparison.Ordinal);
        }
        finally
        {
            await Task.WhenAll(StopTestServiceAsync(defaultDaemon, defaultPipe),
                StopTestServiceAsync(selectedDaemon, selectedPipe));
        }
    }

    private static async Task StopTestServiceAsync(Process daemon, string pipeName)
    {
        try
        {
            using var client = new ServiceClient(pipeName, requestTimeout: TimeSpan.FromSeconds(2));
            await client.SendAsync(new ServiceRequest { Command = "service.shutdown" });
            await daemon.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            if (!daemon.HasExited)
            {
                daemon.Kill();
                await daemon.WaitForExitAsync();
            }
        }
    }

    [Theory]
    [InlineData("status")]
    [InlineData("stop")]
    public async Task StoppedService_ReportsWithoutWaitingForPipeConnection(string action)
    {
        using var process = StartCli("service", action, "--pipe-name", $"PowerPointMcp_Stopped_{Guid.NewGuid():N}");
        try
        {
            var exit = process.WaitForExitAsync();
            await Task.WhenAny(exit, Task.Delay(TimeSpan.FromSeconds(2)));
            Assert.True(exit.IsCompleted, "An absent service must not require a pipe connection timeout.");
            Assert.Equal(0, process.ExitCode);
        }
        finally
        {
            if (!process.HasExited)
            {
                process.Kill();
                await process.WaitForExitAsync();
            }
        }
    }

    [Fact]
    public async Task ServiceStop_DisposesShutdownConnectionBeforeWaitingForDaemonExit()
    {
        var pipeName = $"PowerPointMcp_StopRegression_{Guid.NewGuid():N}";
        using var daemon = StartCli(
            "service", "run",
            "--pipe-name", pipeName,
            "--idle-timeout-minutes", "5");

        try
        {
            await WaitForDaemonAsync(pipeName);

            using var stop = StartCli("service", "stop", "--pipe-name", pipeName);
            var stopOutput = await stop.StandardOutput.ReadToEndAsync();
            var stopError = await stop.StandardError.ReadToEndAsync();
            await stop.WaitForExitAsync();

            Assert.True(stop.ExitCode == 0, $"{stopOutput}{Environment.NewLine}{stopError}");
            Assert.Contains("Daemon shutdown started.", stopOutput, StringComparison.Ordinal);

            using var exitTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await daemon.WaitForExitAsync(exitTimeout.Token);
            Assert.Equal(0, daemon.ExitCode);
        }
        finally
        {
            if (!daemon.HasExited)
            {
                daemon.Kill(entireProcessTree: true);
                await daemon.WaitForExitAsync();
            }
        }
    }

    [Fact]
    public async Task Service_ShutdownRejectsNewSessionCommands()
    {
        using var service = new PowerPointMcpService();

        var shutdown = await service.ProcessAsync(new ServiceRequest
        {
            Command = "service.shutdown",
            Source = "cli"
        });
        var afterShutdown = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.list",
            Source = "cli"
        });

        Assert.True(shutdown.Success);
        Assert.False(afterShutdown.Success);
        Assert.Equal("ServiceUnavailable", afterShutdown.ErrorCategory);
        Assert.Contains("shutting down", afterShutdown.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    private static async Task WaitForDaemonAsync(string pipeName)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        while (DateTime.UtcNow < deadline)
        {
            using var status = StartCli("service", "status", "--pipe-name", pipeName);
            var output = await status.StandardOutput.ReadToEndAsync();
            await status.WaitForExitAsync();

            if (status.ExitCode == 0 && output.Contains("\"responsive\":true", StringComparison.Ordinal))
            {
                return;
            }

            await Task.Delay(200);
        }

        throw new TimeoutException("Test daemon did not become responsive.");
    }

    private static Process StartCli(params string[] arguments)
        => StartCliWithDefaultPipe(null, arguments);

    private static Process StartCliWithDefaultPipe(string? defaultPipe, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = Path.Combine(AppContext.BaseDirectory, "powerpointcli.exe"),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (defaultPipe != null)
            startInfo.Environment["POWERPOINTMCP_CLI_PIPE"] = defaultPipe;

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        return Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start powerpointcli.exe.");
    }
}
