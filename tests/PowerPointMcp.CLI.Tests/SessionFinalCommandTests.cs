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
        var getExitCode = await Program.Main(
            ["session", "get-final", "missing-session", "--is-final", "true"]);
        Assert.NotEqual(0, getExitCode);

        var setExitCode = await Program.Main(
            ["session", "set-final", "missing-session"]);
        Assert.NotEqual(0, setExitCode);
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
}
