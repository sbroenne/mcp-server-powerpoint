using Sbroenne.PowerPointMcp.CLI.Commands;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class TagCommandContractTests
{
    [Fact]
    public void PresentationTagCommands_ExposeStrictSettings()
    {
        Assert.True(typeof(AsyncCommand<SessionTagSetSettings>).IsAssignableFrom(typeof(SessionSetTagCommand)));
        Assert.True(typeof(AsyncCommand<SessionTagGetSettings>).IsAssignableFrom(typeof(SessionGetTagCommand)));
        Assert.True(typeof(AsyncCommand<SessionIdSettings>).IsAssignableFrom(typeof(SessionListTagsCommand)));
        Assert.True(typeof(AsyncCommand<SessionTagGetSettings>).IsAssignableFrom(typeof(SessionDeleteTagCommand)));

        var settings = new SessionTagSetSettings
        {
            SessionId = "session",
            TagName = "OWNER",
            TagValue = "MiXeD Value"
        };
        Assert.Equal("session", settings.SessionId);
        Assert.Equal("OWNER", settings.TagName);
        Assert.Equal("MiXeD Value", settings.TagValue);
    }
}
