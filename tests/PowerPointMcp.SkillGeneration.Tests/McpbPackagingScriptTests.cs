using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;

namespace Sbroenne.PowerPointMcp.SkillGeneration.Tests;

[Trait("RequiresPowerPoint", "false")]
public sealed class McpbPackagingScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    public void PackagingScript_UsesVerifiedMetadataOnlyStaging()
    {
        var script = File.ReadAllText(Path.Combine(RepoRoot, "mcpb", "Build-McpBundle.ps1"));

        Assert.Contains("McpbPackaging.ps1", script);
        Assert.Contains("ZipFile]::OpenRead", script);
        Assert.Contains("Remove-McpbStagingDirectory", script);
        Assert.DoesNotContain("dotnet publish", script, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("npm install", script, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PackagingScript_CreatesOnlyExpectedMetadata()
    {
        var output = Path.Combine(RepoRoot, "artifacts", $"mcpb-test-{Guid.NewGuid():N}");
        try
        {
            var result = await RunPowerShellAsync(
                Path.Combine(RepoRoot, "mcpb", "Build-McpBundle.ps1"),
                "-Version", "9.8.7",
                "-OutputDir", output);
            Assert.True(result.ExitCode == 0, result.Output);

            var bundle = Path.Combine(output, "powerpoint-mcp-9.8.7.mcpb");
            Assert.True(File.Exists(bundle), result.Output);
            using var zip = ZipFile.OpenRead(bundle);
            Assert.Equal(
                ["CHANGELOG.md", "LICENSE", "README.md", "icon-512.png", "manifest.json"],
                zip.Entries.Select(entry => entry.FullName).Order(StringComparer.Ordinal).ToArray());

            var manifestEntry = Assert.Single(zip.Entries, entry => entry.FullName == "manifest.json");
            using var stream = manifestEntry.Open();
            using var manifest = await JsonDocument.ParseAsync(stream);
            Assert.Equal("9.8.7", manifest.RootElement.GetProperty("version").GetString());
            Assert.Equal("npx", manifest.RootElement.GetProperty("server").GetProperty("mcp_config")
                .GetProperty("command").GetString());
        }
        finally
        {
            if (Directory.Exists(output))
                Directory.Delete(output, recursive: true);
        }
    }

    private static async Task<(int ExitCode, string Output)> RunPowerShellAsync(
        string script,
        params string[] arguments)
    {
        var startInfo = new ProcessStartInfo("pwsh")
        {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(script);
        foreach (var argument in arguments)
            startInfo.ArgumentList.Add(argument);

        using var process = Process.Start(startInfo)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromMinutes(2));
        return (process.ExitCode, $"{await stdout}{await stderr}");
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.PowerPointMcp.slnx")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("Repository root not found.");
    }
}
