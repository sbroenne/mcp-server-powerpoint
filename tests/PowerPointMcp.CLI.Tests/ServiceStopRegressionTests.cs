using System.Diagnostics;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class ServiceStopRegressionTests
{
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
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = Path.Combine(AppContext.BaseDirectory, "powerpointcli.exe"),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        return Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start powerpointcli.exe.");
    }
}
