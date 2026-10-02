using System.Diagnostics;
using System.Text.Json;

namespace Sbroenne.PowerPointMcp.SkillGeneration.Tests;

public sealed class PluginBootstrapBuildTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    public void McpPlugin_UsesPublicNpmPackage()
    {
        var path = Path.Combine(RepoRoot, ".github", "plugins", "powerpoint-mcp", "mcp.json");
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var server = document.RootElement.GetProperty("mcpServers").GetProperty("powerpoint-mcp");

        Assert.Equal("npx", server.GetProperty("command").GetString());
        Assert.Equal(
            ["-y", "@sbroenne/mcp-server-powerpoint@latest"],
            server.GetProperty("args").EnumerateArray().Select(value => value.GetString()!).ToArray());
    }

    [Fact]
    public void PluginSources_DoNotContainRetiredDownloadersOrGlobalInstallers()
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins");
        Assert.Empty(Directory.GetFiles(pluginRoot, "download.ps1", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(pluginRoot, "install-global.ps1", SearchOption.AllDirectories));
        Assert.False(File.Exists(Path.Combine(RepoRoot, "scripts", "Build-BootstrapScripts.ps1")));
    }

    [Fact]
    public void CliPlugin_UsesArgumentSafeNpxWrapper()
    {
        var wrapper = File.ReadAllText(Path.Combine(
            RepoRoot,
            ".github",
            "plugins",
            "powerpoint-cli",
            "bin",
            "start-cli.ps1"));

        Assert.Contains("@sbroenne/pptcli@latest", wrapper, StringComparison.Ordinal);
        Assert.Contains("ConvertTo-NativeArgument", wrapper, StringComparison.Ordinal);
        Assert.Contains("npx-cli.js", wrapper, StringComparison.Ordinal);
        Assert.DoesNotContain("download.ps1", wrapper, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildPlugins_ProducesNpxOnlyPlugins()
    {
        var output = Path.Combine(Path.GetTempPath(), $"PowerPointMcpPlugins-{Guid.NewGuid():N}");
        try
        {
            var result = RunPowerShell(
                Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1"),
                "-Version",
                "9.8.7",
                "-OutputDir",
                output);

            Assert.Equal(0, result.ExitCode);
            Assert.True(File.Exists(Path.Combine(output, "powerpoint-mcp", "mcp.json")));
            Assert.True(File.Exists(Path.Combine(output, "powerpoint-cli", "bin", "start-cli.ps1")));
            Assert.Empty(Directory.GetFiles(output, "download.ps1", SearchOption.AllDirectories));
            Assert.Empty(Directory.GetFiles(output, "install-global.ps1", SearchOption.AllDirectories));
        }
        finally
        {
            if (Directory.Exists(output))
            {
                Directory.Delete(output, recursive: true);
            }
        }
    }

    private static (int ExitCode, string Output) RunPowerShell(string script, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(script);
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo)!;
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, stdout + stderr);
    }

    private static string FindRepoRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null && !File.Exists(Path.Combine(current.FullName, "Sbroenne.PowerPointMcp.slnx")))
        {
            current = current.Parent;
        }

        return current?.FullName ?? throw new DirectoryNotFoundException("Repository root not found.");
    }
}
