using System.Diagnostics;
using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class SessionCommandContractTests
{
    [Fact]
    public async Task SessionHelp_ListsSaveAsAndSaveCopyAs()
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = Path.Combine(AppContext.BaseDirectory, "powerpointcli.exe"),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("session");
        startInfo.ArgumentList.Add("--help");

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);

        string output = await process.StandardOutput.ReadToEndAsync();
        string error = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        Assert.True(process.ExitCode == 0, error);
        Assert.Contains("save-as", output, StringComparison.Ordinal);
        Assert.Contains("save-copy-as", output, StringComparison.Ordinal);
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
