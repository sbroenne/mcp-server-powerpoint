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
    public void CoreTestProject_IsExplicitlyMarkedForTestDiscovery()
    {
        var project = XDocument.Load(Path.Combine(
            RepoRoot,
            "tests",
            "PowerPointMcp.Core.Tests",
            "PowerPointMcp.Core.Tests.csproj"));

        Assert.Equal("true", project.Root!
            .Elements("PropertyGroup")
            .Elements("IsTestProject")
            .SingleOrDefault()?.Value);
    }

    [Fact]
    public void ValidationGates_RunReleasePackagingTests()
    {
        const string testProject = "PowerPointMcp.SkillGeneration.Tests";

        Assert.Contains(testProject, File.ReadAllText(CiWorkflow), StringComparison.Ordinal);
        Assert.Contains(testProject, File.ReadAllText(PreCommitScript), StringComparison.Ordinal);
    }

    [Fact]
    public void ComLeakAudit_RejectsUnrelatedRelease()
    {
        var result = RunComLeakAudit("""
            class Commands
            {
                void Read(dynamic source)
                {
                    dynamic released = source.First;
                    dynamic leaked = source.Second;
                    ComUtilities.Release(ref released);
                }
            }
            """);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("leaked", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("dynamic item = source.Child; ComUtilities.Release(ref item);", 0, 1)]
    [InlineData("dynamic /* optional */ ? item = source.Child;", 1, 1)]
    [InlineData("dynamic item = source.Child; ComUtilities /* cleanup */ .Release(ref item);", 0, 1)]
    [InlineData("dynamic @item = source.Child; ComUtilities.Release(ref @item!);", 0, 1)]
    [InlineData("dynamic? item = null; try { item = source.Child; } finally { ComUtilities.Release(ref item!); }", 0, 1)]
    [InlineData("dynamic? item = null; item = source.Child;", 1, 1)]
    [InlineData("dynamic borrowed = ctx.Presentation;", 0, 0)]
    [InlineData("dynamic borrowed = ctx.App;", 0, 0)]
    [InlineData("dynamic dispatch = source;", 0, 0)]
    [InlineData("dynamic dispatch = (dynamic)(source!);", 0, 0)]
    [InlineData("dynamic? dispatch = null; dispatch = source;", 0, 0)]
    [InlineData("dynamic item = (dynamic)(source.Child!);", 1, 1)]
    [InlineData("dynamic? item = null;", 0, 0)]
    [InlineData("dynamic item = source.Child; /* ComUtilities.Release(ref item); */", 1, 1)]
    [InlineData("dynamic item = source.Child; var text = \"ComUtilities.Release(ref item);\";", 1, 1)]
    [InlineData("var text = \"dynamic leaked = source.Child;\";", 0, 0)]
    [InlineData("dynamic item = source.Child; ComUtilities.Release(ref Item);", 1, 1)]
    [InlineData("dynamic item = source.Child; void Cleanup() { ComUtilities.Release(ref item); }", 1, 1)]
    public void ComLeakAudit_RecognizesSupportedSyntax(string body, int exitCode, int acquisitions)
    {
        var result = RunComLeakAudit($"class Commands {{ void Read(dynamic source) {{ {body} }} }}");

        Assert.True(result.ExitCode == exitCode, result.Output);
        Assert.Contains($"Checked {acquisitions} dynamic acquisition variables", result.Output, StringComparison.Ordinal);
        Assert.Contains("typed PIA ownership, control flow, and release-in-finally are not verified", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("void Other() { dynamic item = source.Child; ComUtilities.Release(ref item); }", "void Read() { dynamic item = source.Child; }")]
    [InlineData("void Read() { { dynamic item = source.Child; ComUtilities.Release(ref item); }", "{ dynamic item = source.Child; } }")]
    public void ComLeakAudit_RejectsReleaseFromAnotherScope(string first, string second)
    {
        var result = RunComLeakAudit($"class Commands {{ {first} {second} }}");

        Assert.True(result.ExitCode == 1, result.Output);
        Assert.Contains("Checked 2 dynamic acquisition variables; 1 missing releases", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("obj/Commands.cs")]
    [InlineData("bin/Commands.cs")]
    [InlineData("Commands.g.cs")]
    public void ComLeakAudit_ExcludesGeneratedFiles(string generatedPath)
    {
        var result = RunComLeakAudit("class Commands { }", generatedPath: generatedPath);

        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("Scanned 1 source files", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true, "No C# source files found")]
    [InlineData(false, "Source directory not found")]
    public void ComLeakAudit_RejectsBrokenSourceDiscovery(bool createSourceDirectory, string message)
    {
        var result = RunComLeakAudit(null, createSourceDirectory);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(message, result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public void ComLeakAudit_RejectsUnparseableSource()
    {
        var result = RunComLeakAudit("class Commands { void Read( {");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("Cannot parse", result.Output, StringComparison.Ordinal);
    }

    private static ProcessResult RunComLeakAudit(string? source, bool createSourceDirectory = true, string? generatedPath = null)
    {
        using var temp = new TemporaryDirectory();
        var root = Path.Combine(temp.Path, "audit fixture");
        var scripts = Path.Combine(root, "scripts");
        var sources = Path.Combine(root, "src");
        Directory.CreateDirectory(scripts);
        if (createSourceDirectory)
        {
            Directory.CreateDirectory(sources);
        }
        var script = Path.Combine(scripts, "check-com-leaks.ps1");
        File.Copy(Path.Combine(RepoRoot, "scripts", "check-com-leaks.ps1"), script);
        if (source != null)
        {
            File.WriteAllText(Path.Combine(sources, "Commands.cs"), source);
        }
        if (generatedPath != null)
        {
            var generatedFile = Path.Combine(sources, generatedPath);
            Directory.CreateDirectory(Path.GetDirectoryName(generatedFile)!);
            File.WriteAllText(generatedFile, "class Generated { void Read() { dynamic leaked = source.Child; } }");
        }

        return RunPowerShellRaw(script);
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

    [Theory]
    [InlineData("")]
    [InlineData("vscode-extension")]
    [InlineData("videos/powerpoint-mcp-intro")]
    public void NpmLockfiles_ProjectConfigOmitsUrlsWithoutOverridingPackageSources(string project)
    {
        var config = Path.Combine(RepoRoot, project, ".npmrc");
        Assert.True(File.Exists(config), $"Missing npm project configuration: {project}");
        Assert.Equal("omit-lockfile-registry-resolved=true", File.ReadAllText(config).Trim());
    }

    [Theory]
    [InlineData("""{"lockfileVersion":3,"packages":{"":{"name":"sample"},"node_modules/sample":{"version":"1.0.0","integrity":"sha512-example"}}}""", true)]
    [InlineData("""{"lockfileVersion":3,"packages":{"node_modules/sample":{"resolved":"https://registry.npmjs.org/sample/-/sample-1.0.0.tgz"}}}""", false)]
    [InlineData("""{"lockfileVersion":3,"packages":{"node_modules/sample":{"resolved":"https://mirror.example.test/sample.tgz"}}}""", false)]
    [InlineData("""{"lockfileVersion":1,"dependencies":{"sample":{"dependencies":{"nested":{"resolved":"https://mirror.example.test/nested.tgz"}}}}}""", false)]
    [InlineData("""{"lockfileVersion":3,"packages":{"node_modules/local":{"resolved":"file:../local"},"node_modules/sample":{"homepage":"https://example.test"}}}""", true)]
    [InlineData("{invalid-json", false)]
    public void NpmLockfiles_GuardAcceptsPortableFilesAndRejectsDownloadUrls(string content, bool succeeds)
    {
        using var temp = new TemporaryDirectory();
        var path = Path.Combine(temp.Path, "package-lock.json");
        File.WriteAllText(path, content);

        var result = RunPowerShellRaw(
            Path.Combine(RepoRoot, "scripts", "check-npm-lockfiles.ps1"),
            "-LockfilePath", path);

        Assert.Equal(succeeds, result.ExitCode == 0);
        Assert.Contains(succeeds ? "passed" : "BLOCKED", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("https://", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public void NpmLockfiles_ValidationGatesRunGuard()
    {
        Assert.Contains("scripts/check-npm-lockfiles.ps1", File.ReadAllText(CiWorkflow), StringComparison.Ordinal);
        Assert.Contains("scripts\\check-npm-lockfiles.ps1", File.ReadAllText(PreCommitScript), StringComparison.Ordinal);
    }

    [Fact]
    public void NpmLockfiles_GuardDiscoversNestedProjectsAndChecksStagedContent()
    {
        using var temp = new TemporaryDirectory();
        const string portable = """{"lockfileVersion":3,"packages":{}}""";
        const string nonportable = """{"lockfileVersion":3,"packages":{"node_modules/sample":{"resolved":"https://mirror.example.test/sample.tgz"}}}""";
        var nested = Path.Combine(temp.Path, "nested project");
        Directory.CreateDirectory(nested);
        var lockfile = Path.Combine(nested, "package-lock.json");
        File.WriteAllText(lockfile, nonportable);
        File.WriteAllText(Path.Combine(temp.Path, "npm-shrinkwrap.json"), portable);
        File.WriteAllText(Path.Combine(temp.Path, "package-lock.json"), nonportable);
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "init", "--quiet").ExitCode);
        Assert.Equal(0, RunProcessRaw(
            "git", "-C", temp.Path, "add", "--", "nested project/package-lock.json", "npm-shrinkwrap.json").ExitCode);
        var script = Path.Combine(RepoRoot, "scripts", "check-npm-lockfiles.ps1");

        var workingResult = RunPowerShellRaw(script, "-RepoRoot", temp.Path);
        Assert.NotEqual(0, workingResult.ExitCode);
        Assert.Contains("BLOCKED", workingResult.Output, StringComparison.Ordinal);

        File.WriteAllText(lockfile, portable);
        RunPowerShell(script, "-RepoRoot", temp.Path);
        var stagedResult = RunPowerShellRaw(script, "-RepoRoot", temp.Path, "-Staged");
        Assert.NotEqual(0, stagedResult.ExitCode);
        Assert.Contains("BLOCKED", stagedResult.Output, StringComparison.Ordinal);

        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "add", "--", "nested project/package-lock.json").ExitCode);
        RunPowerShell(script, "-RepoRoot", temp.Path, "-Staged");
    }

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
        => RunProcessRaw("pwsh", ["-NoProfile", "-File", script, .. arguments]);

    private static ProcessResult RunProcessRaw(string executable, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        // Commit hooks export repository paths that must not leak into temporary test repositories.
        foreach (var variable in new[]
        {
            "GIT_DIR", "GIT_WORK_TREE", "GIT_IMPLICIT_WORK_TREE", "GIT_INDEX_FILE",
            "GIT_COMMON_DIR", "GIT_OBJECT_DIRECTORY", "GIT_ALTERNATE_OBJECT_DIRECTORIES",
            "GIT_PREFIX", "GIT_GRAFT_FILE", "GIT_SHALLOW_FILE",
        })
        {
            startInfo.Environment.Remove(variable);
        }
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Failed to start {executable}.");
        var standardOutput = process.StandardOutput.ReadToEndAsync();
        var standardError = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"Process timed out: {executable}");
        }

        return new ProcessResult(
            process.ExitCode,
            $"{standardOutput.GetAwaiter().GetResult()}{Environment.NewLine}{standardError.GetAwaiter().GetResult()}");
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
                // Git marks loose objects read-only, including in isolated test repositories.
                foreach (var file in Directory.EnumerateFiles(Path, "*", SearchOption.AllDirectories))
                {
                    File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                }
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
