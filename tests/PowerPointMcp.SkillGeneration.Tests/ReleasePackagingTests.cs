using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

namespace Sbroenne.PowerPointMcp.SkillGeneration.Tests;

public sealed class ReleasePackagingTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string ReleaseWorkflow = Path.Combine(
        RepoRoot,
        ".github",
        "workflows",
        "release.yml");
    private static readonly string CiWorkflow = Path.Combine(
        RepoRoot,
        ".github",
        "workflows",
        "ci.yml");
    private static readonly string PreCommitScript = Path.Combine(
        RepoRoot,
        "scripts",
        "pre-commit.ps1");

    [Fact]
    public void UpdateReleaseVersionMetadata_StampsEveryPersistentVersion()
    {
        using var temp = new TemporaryDirectory();
        foreach (var relativePath in MetadataPaths)
        {
            var source = Path.Combine(RepoRoot, relativePath);
            var destination = Path.Combine(temp.Path, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(source, destination);
        }

        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "Update-ReleaseVersionMetadata.ps1"),
            "-Version", "9.8.7",
            "-RepoRoot", temp.Path);

        AssertJsonVersion(Path.Combine(temp.Path, "package.json"), "9.8.7");
        AssertPackageLockVersions(Path.Combine(temp.Path, "package-lock.json"), "9.8.7");
        AssertJsonVersion(Path.Combine(temp.Path, "mcpb", "manifest.json"), "9.8.7");
        AssertJsonVersion(Path.Combine(temp.Path, "vscode-extension", "package.json"), "9.8.7");
        AssertPackageLockVersions(
            Path.Combine(temp.Path, "vscode-extension", "package-lock.json"),
            "9.8.7");

        var props = XDocument.Load(Path.Combine(temp.Path, "Directory.Build.props"));
        Assert.Equal("9.8.7", props.Descendants("Version").Single().Value);
        Assert.Equal("9.8.7.0", props.Descendants("AssemblyVersion").Single().Value);
        Assert.Equal("9.8.7.0", props.Descendants("FileVersion").Single().Value);

        using var server = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(temp.Path, "src", "PowerPointMcp.McpServer", ".mcp", "server.json")));
        Assert.Equal("9.8.7", server.RootElement.GetProperty("version").GetString());
        Assert.Equal(
            "9.8.7",
            server.RootElement.GetProperty("packages")[0].GetProperty("version").GetString());
    }

    [Fact]
    public void BuildAgentSkills_CreatesBothVersionedSkillsFromGeneratedReferences()
    {
        using var temp = new TemporaryDirectory();
        var cliPath = Path.Combine(
            RepoRoot,
            "src",
            "PowerPointMcp.CLI",
            "bin",
            "Release",
            "net10.0-windows",
            "powerpointcli.exe");

        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1"),
            "-Version", "9.8.7",
            "-OutputDir", temp.Path,
            "-CliPath", cliPath);

        Assert.False(File.Exists(Path.Combine(RepoRoot, "skills", "powerpoint-mcp", "VERSION")));
        var zipPath = Assert.Single(Directory.GetFiles(temp.Path, "*.zip"));
        using var archive = ZipFile.OpenRead(zipPath);

        AssertEntryText(archive, "skills/powerpoint-mcp/VERSION", "9.8.7");
        AssertEntryText(archive, "skills/powerpoint-cli/VERSION", "9.8.7");
        var cliReference = ReadEntry(archive, "skills/powerpoint-cli/references/cli-commands.md");
        Assert.Contains("pptcli session", cliReference, StringComparison.Ordinal);
        Assert.Contains("pptcli service stop", cliReference, StringComparison.Ordinal);

        using var manifest = JsonDocument.Parse(
            File.ReadAllText(Path.Combine(temp.Path, "manifest.json")));
        Assert.Equal("9.8.7", manifest.RootElement.GetProperty("version").GetString());
        Assert.Equal(2, manifest.RootElement.GetProperty("skills").GetArrayLength());
    }

    [Fact]
    public void ReleaseWorkflow_UsesCanonicalScriptsChecksumsAndStrictRegistryPublishing()
    {
        var workflow = File.ReadAllText(ReleaseWorkflow);

        Assert.Contains(
            "./scripts/Update-ReleaseVersionMetadata.ps1 -Version $env:VERSION",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "./scripts/Update-McpRegistryMetadata.ps1",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "./scripts/Build-AgentSkills.ps1 -Version $env:VERSION",
            workflow,
            StringComparison.Ordinal);
        Assert.DoesNotContain("$serverContent = $serverContent -replace", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("buildDate", workflow, StringComparison.Ordinal);

        Assert.Contains("name: Standalone Checksums", workflow, StringComparison.Ordinal);
        Assert.Contains("sha256sum *.zip", workflow, StringComparison.Ordinal);
        Assert.Contains("artifacts/standalone-checksums/SHA256SUMS", workflow, StringComparison.Ordinal);

        var registryStepStart = workflow.IndexOf("- name: Publish to MCP Registry", StringComparison.Ordinal);
        Assert.True(registryStepStart >= 0);
        var nextJobStart = workflow.IndexOf(
            "  # =============================================================================",
            registryStepStart,
            StringComparison.Ordinal);
        Assert.True(nextJobStart > registryStepStart);
        Assert.DoesNotContain(
            "continue-on-error",
            workflow[registryStepStart..nextJobStart],
            StringComparison.Ordinal);
    }

    [Fact]
    public void ValidationGates_RunReleasePackagingTests()
    {
        const string testProject = "PowerPointMcp.SkillGeneration.Tests";

        Assert.Contains(testProject, File.ReadAllText(CiWorkflow), StringComparison.Ordinal);
        Assert.Contains(testProject, File.ReadAllText(PreCommitScript), StringComparison.Ordinal);
    }

    [Fact]
    public void RegistryMetadataScript_RejectsMissingServerPackage()
    {
        using var temp = new TemporaryDirectory();
        var serverJson = Path.Combine(temp.Path, "server.json");
        File.WriteAllText(serverJson, """{"version":"1.0.0","packages":[]}""");

        var result = RunPowerShellRaw(
            Path.Combine(RepoRoot, "scripts", "Update-McpRegistryMetadata.ps1"),
            "-ServerJsonPath",
            serverJson,
            "-Version",
            "9.8.7");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("exactly one", result.Output, StringComparison.OrdinalIgnoreCase);
    }

    private static readonly string[] MetadataPaths =
    [
        "package.json",
        "package-lock.json",
        "Directory.Build.props",
        Path.Combine("mcpb", "manifest.json"),
        Path.Combine("vscode-extension", "package.json"),
        Path.Combine("vscode-extension", "package-lock.json"),
        Path.Combine("src", "PowerPointMcp.McpServer", ".mcp", "server.json"),
    ];

    private static void AssertJsonVersion(string path, string expected)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.Equal(expected, document.RootElement.GetProperty("version").GetString());
    }

    private static void AssertPackageLockVersions(string path, string expected)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.Equal(expected, document.RootElement.GetProperty("version").GetString());
        Assert.Equal(
            expected,
            document.RootElement.GetProperty("packages").GetProperty("").GetProperty("version").GetString());
    }

    private static void AssertEntryText(ZipArchive archive, string path, string expected)
    {
        Assert.Equal(expected, ReadEntry(archive, path));
    }

    private static string ReadEntry(ZipArchive archive, string path)
    {
        var entry = archive.Entries.SingleOrDefault(candidate =>
            string.Equals(
                candidate.FullName.Replace('\\', '/'),
                path,
                StringComparison.Ordinal));
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static void RunPowerShell(string script, params string[] arguments)
    {
        var result = RunPowerShellRaw(script, arguments);
        Assert.True(
            result.ExitCode == 0,
            $"PowerShell failed.{Environment.NewLine}{result.Output}");
    }

    private static ProcessResult RunPowerShellRaw(string script, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(script);
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start PowerShell.");
        var standardOutput = process.StandardOutput.ReadToEnd();
        var standardError = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(120_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"PowerShell timed out: {script}");
        }

        return new ProcessResult(
            process.ExitCode,
            $"{standardOutput}{Environment.NewLine}{standardError}");
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.PowerPointMcp.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName
            ?? throw new InvalidOperationException("Could not locate the repository root.");
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        internal TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PowerPointMcp.Tests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
