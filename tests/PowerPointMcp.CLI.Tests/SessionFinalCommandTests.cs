using System.Diagnostics;
using System.Text.Json;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class SessionFinalCommandTests
{
    [Theory]
    [InlineData("get-final")]
    [InlineData("set-final")]
    public async Task SessionFinalCommand_IsRegistered(string action)
    {
        var exitCode = await Program.Main(["session", action, "--help"]);

        Assert.Equal(0, exitCode);
    }

    [Fact]
    public async Task SessionFinalCommands_RejectMissingOrInapplicableArguments()
    {
        string pipeName = $"PowerPointMcp_FinalArguments_{Guid.NewGuid():N}";
        using var service = new PowerPointMcpService();
        var running = service.RunAsync(pipeName, TimeSpan.FromMinutes(5));
        try
        {
            var getExitCode = await RunCliAsync(pipeName,
                "session", "get-final", "missing-session", "--is-final", "true");
            Assert.NotEqual(0, getExitCode);

            var setExitCode = await RunCliAsync(pipeName,
                "session", "set-final", "missing-session");
            Assert.NotEqual(0, setExitCode);
        }
        finally
        {
            service.RequestShutdown();
            await running.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact]
    public async Task FinalServiceActions_RejectUnknownSessionsWithoutStartingPowerPoint()
    {
        using var service = new PowerPointMcpService();

        var getResponse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.get-final",
            SessionId = "missing-session",
            Source = "cli"
        });
        Assert.False(getResponse.Success);
        Assert.Contains("not found", getResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setResponse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.set-final",
            SessionId = "missing-session",
            Args = JsonSerializer.Serialize(new { isFinal = true }, ServiceProtocol.JsonOptions),
            Source = "cli"
        });
        Assert.False(setResponse.Success);
        Assert.Contains("not found", setResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FinalServiceActions_RejectUnknownParameters()
    {
        using var service = new PowerPointMcpService();

        var getResponse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.get-final",
            SessionId = "missing-session",
            Args = JsonSerializer.Serialize(new { isFinal = true }, ServiceProtocol.JsonOptions),
            Source = "cli"
        });
        Assert.False(getResponse.Success);
        Assert.Contains("Unknown parameter", getResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setResponse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.set-final",
            SessionId = "missing-session",
            Args = JsonSerializer.Serialize(new { isFinal = true, value = "unexpected" }, ServiceProtocol.JsonOptions),
            Source = "cli"
        });
        Assert.False(setResponse.Success);
        Assert.Contains("Unknown parameter", setResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<int> RunCliAsync(string pipeName, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = Path.Combine(AppContext.BaseDirectory, "powerpointcli.exe"),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.Environment["POWERPOINTMCP_CLI_PIPE"] = pipeName;
        foreach (string argument in arguments)
            startInfo.ArgumentList.Add(argument);

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start the test CLI.");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            await Task.WhenAll(output, error);
            return process.ExitCode;
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
}
