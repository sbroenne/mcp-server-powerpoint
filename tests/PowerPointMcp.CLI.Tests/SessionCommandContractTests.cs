using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class SessionCommandContractTests
{
    [Fact]
    public async Task SessionHelp_ListsSaveAsAndSaveCopyAs()
    {
        var originalOut = Console.Out;
        using var output = new StringWriter();
        Console.SetOut(output);
        try
        {
            int exitCode = await Program.Main(["session", "--help"]);

            Assert.Equal(0, exitCode);
            Assert.Contains("save-as", output.ToString(), StringComparison.Ordinal);
            Assert.Contains("save-copy-as", output.ToString(), StringComparison.Ordinal);
        }
        finally
        {
            Console.SetOut(originalOut);
        }
    }

    [Fact]
    public async Task Service_SaveAsContract_DeserializesFormatAndRoutesToSessionLookup()
    {
        using var service = new PowerPointMcpService();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save-as",
            SessionId = "missing-session",
            Args = JsonSerializer.Serialize(
                new
                {
                    targetPath = @"C:\Temp\saved.pptx",
                    format = PresentationSaveFormat.Pptx,
                    overwrite = true
                },
                ServiceProtocol.JsonOptions),
            Source = "cli"
        });

        Assert.False(response.Success);
        Assert.Contains("not found", response.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("Unknown session action", response.ErrorMessage, StringComparison.Ordinal);
    }
}
