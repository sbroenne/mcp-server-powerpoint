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
    public void CoreInterfaceGuard_HandlesParenthesesInParameterDescriptionsAndStillDetectsOrphans()
    {
        using var temp = new TemporaryDirectory();
        var scripts = Path.Combine(temp.Path, "scripts");
        var domain = Path.Combine(temp.Path, "src", "PowerPointMcp.Core", "Sample");
        Directory.CreateDirectory(scripts);
        Directory.CreateDirectory(domain);
        var script = Path.Combine(scripts, "check-core-interface-completeness.ps1");
        File.Copy(Path.Combine(RepoRoot, "scripts", "check-core-interface-completeness.ps1"), script);
        File.WriteAllText(Path.Combine(domain, "ISampleCommands.cs"), """
            public interface ISampleCommands
            {
                string Inspect(
                    [System.ComponentModel.Description("Maximum (1-100; default 20).")]
                    int maxSlides = 20);
            }
            """);
        var implementation = Path.Combine(domain, "SampleCommands.cs");
        File.WriteAllText(implementation, """
            public class SampleCommands : ISampleCommands
            {
                public string Inspect(int maxSlides = 20) { return ""; }
            }
            """);
        RunPowerShell(script);

        File.WriteAllText(implementation, """
            public class SampleCommands : ISampleCommands
            {
                public string Inspect(int maxSlides = 20) { return ""; }
                public string Missing() { return ""; }
            }
            """);
        var result = RunPowerShellRaw(script);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("SampleCommands.Missing", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("SampleCommands.Inspect", result.Output, StringComparison.Ordinal);
    }

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
    public void CliSkill_RemainsCompactAndIsNotRegeneratedDuringBuild()
    {
        var project = XDocument.Load(Path.Combine(
            RepoRoot, "src", "PowerPointMcp.CLI", "PowerPointMcp.CLI.csproj"));
        Assert.DoesNotContain(
            project.Descendants("Target"),
            target => (string?)target.Attribute("Name") == "GenerateCliSkill");
        foreach (var skillName in new[] { "powerpoint-cli", "powerpoint-mcp" })
        {
            var skill = File.ReadAllText(Path.Combine(RepoRoot, "skills", skillName, "SKILL.md"));
            Assert.True(skill.Length < 4000, $"{skillName} should remain a compact entry skill.");
        }
    }

    [Fact]
    public void BuildAgentSkills_PackagesThreeFocusedVersionedSkills()
    {
        using var temp = new TemporaryDirectory();

        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1"),
            "-Version", "9.8.7",
            "-OutputDir", temp.Path);

        Assert.False(File.Exists(Path.Combine(RepoRoot, "skills", "powerpoint-mcp", "VERSION")));
        var zipPath = Assert.Single(Directory.GetFiles(temp.Path, "*.zip"));
        using var archive = ZipFile.OpenRead(zipPath);

        foreach (var skillName in new[] { "powerpoint-mcp", "powerpoint-cli", "powerpoint-deck-design" })
        {
            AssertEntryText(archive, $"skills/{skillName}/VERSION", "9.8.7");
            Assert.Contains(
                $"name: {skillName}",
                ReadEntry(archive, $"skills/{skillName}/SKILL.md"),
                StringComparison.Ordinal);
        }
        var expectedEntries = new[]
        {
            "README.md",
            "skills/powerpoint-cli/SKILL.md",
            "skills/powerpoint-cli/VERSION",
            "skills/powerpoint-deck-design/SKILL.md",
            "skills/powerpoint-deck-design/VERSION",
            "skills/powerpoint-mcp/SKILL.md",
            "skills/powerpoint-mcp/VERSION"
        };
        Assert.Equal(
            expectedEntries.OrderBy(path => path, StringComparer.Ordinal),
            archive.Entries.Select(entry => entry.FullName).OrderBy(path => path, StringComparer.Ordinal));

        using var manifest = JsonDocument.Parse(
            File.ReadAllText(Path.Combine(temp.Path, "manifest.json")));
        Assert.Equal("9.8.7", manifest.RootElement.GetProperty("version").GetString());
        Assert.Equal(3, manifest.RootElement.GetProperty("skills").GetArrayLength());
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

        Assert.Contains("resume_release:", workflow, StringComparison.Ordinal);
        Assert.Contains(
            "./scripts/Assert-ReleaseTagState.ps1",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "source_ref=$TAG",
            workflow,
            StringComparison.Ordinal);
        Assert.Equal(
            10,
            System.Text.RegularExpressions.Regex.Count(
                workflow,
                System.Text.RegularExpressions.Regex.Escape(
                    "ref: ${{ needs.version.outputs.source_ref }}")));
        Assert.Contains(
            "needs: [version, build-cli, build-mcp-server]",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "./scripts/Publish-NpmPackage.ps1",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "gh release upload \"$TAG\" $ARTIFACTS --clobber",
            workflow,
            StringComparison.Ordinal);
        Assert.DoesNotContain("gh release edit", workflow, StringComparison.Ordinal);
        Assert.Contains(
            "Release documentation PR #$EXISTING_PR is already merged.",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "gh pr merge \"$EXISTING_PR\" --squash --delete-branch --auto",
            workflow,
            StringComparison.Ordinal);
        Assert.Contains(
            "ref: ${{ needs.version.outputs.source_ref }}",
            workflow,
            StringComparison.Ordinal);

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
    public void ReleaseTagState_RequiresResumeExactlyWhenTagExists()
    {
        using var temp = new TemporaryDirectory();
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "init").ExitCode);
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "config", "user.name", "Release Test").ExitCode);
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "config", "user.email", "release@example.invalid").ExitCode);
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "commit", "--allow-empty", "-m", "initial").ExitCode);
        Assert.Equal(0, RunProcessRaw("git", "-C", temp.Path, "tag", "v1.2.3").ExitCode);

        var script = Path.Combine(RepoRoot, "scripts", "Assert-ReleaseTagState.ps1");
        Assert.Equal(
            0,
            RunPowerShellRaw(
                script,
                "-Version", "1.2.3",
                "-RepositoryRoot", temp.Path,
                "-ResumeRelease").ExitCode);
        Assert.NotEqual(
            0,
            RunPowerShellRaw(
                script,
                "-Version", "1.2.3",
                "-RepositoryRoot", temp.Path).ExitCode);
        Assert.NotEqual(
            0,
            RunPowerShellRaw(
                script,
                "-Version", "2.0.0",
                "-RepositoryRoot", temp.Path,
                "-ResumeRelease").ExitCode);
        Assert.Equal(
            0,
            RunPowerShellRaw(
                script,
                "-Version", "2.0.0",
                "-RepositoryRoot", temp.Path).ExitCode);
    }

    [Fact]
    public void PublishNpmPackage_SkipsExistingVersionAndPublishesMissingVersion()
    {
        using var temp = new TemporaryDirectory();
        var publishScript = Path.Combine(RepoRoot, "scripts", "Publish-NpmPackage.ps1");
        var logPath = Path.Combine(temp.Path, "npm.log");
        var fakeNpm = Path.Combine(temp.Path, "fake-npm.ps1");
        File.WriteAllText(
            fakeNpm,
            """
            param([Parameter(ValueFromRemainingArguments = $true)][string[]] $Arguments)
            if ($Arguments[0] -eq 'view') {
                if ($Arguments[1] -match 'existing') {
                    Write-Output '1.2.3'
                    exit 0
                }
                exit 1
            }
            if ($Arguments[0] -eq 'publish') {
                if ($Arguments[1] -match 'fail') {
                    exit 3
                }
                Add-Content -Path $env:FAKE_NPM_LOG -Value ($Arguments -join ' ')
                exit 0
            }
            exit 2
            """);

        var previousLog = Environment.GetEnvironmentVariable("FAKE_NPM_LOG");
        Environment.SetEnvironmentVariable("FAKE_NPM_LOG", logPath);
        try
        {
            RunPowerShell(
                publishScript,
                "-PackageName", "@sbroenne/existing",
                "-Version", "1.2.3",
                "-PackageTarball", Path.Combine(temp.Path, "existing.tgz"),
                "-NpmCommand", fakeNpm);
            Assert.False(File.Exists(logPath));

            RunPowerShell(
                publishScript,
                "-PackageName", "@sbroenne/missing",
                "-Version", "1.2.3",
                "-PackageTarball", Path.Combine(temp.Path, "missing.tgz"),
                "-NpmCommand", fakeNpm);
            Assert.Contains("publish", File.ReadAllText(logPath), StringComparison.Ordinal);

            Assert.NotEqual(
                0,
                RunPowerShellRaw(
                    publishScript,
                    "-PackageName", "@sbroenne/fail",
                    "-Version", "1.2.3",
                    "-PackageTarball", Path.Combine(temp.Path, "fail.tgz"),
                    "-NpmCommand", fakeNpm).ExitCode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("FAKE_NPM_LOG", previousLog);
        }
    }

    [Fact]
    public void DocumentationCounts_UpdateValidateAndAllowStaleAdvertisedCounts()
    {
        using var temp = new TemporaryDirectory();
        var expected = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var relativePath in DocumentationCountPaths)
        {
            var source = Path.Combine(RepoRoot, relativePath);
            var destination = Path.Combine(temp.Path, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            expected[relativePath] = File.ReadAllText(source);
            File.Copy(source, destination);
        }

        CorruptOnce(
            Path.Combine(temp.Path, "README.md"),
            @"\d+ MCP tools with \d+ operations",
            "1 MCP tools with 2 operations");
        CorruptOnce(
            Path.Combine(temp.Path, "README.md"),
            @"\*\*Presentation\*\* \(\d+ ops\)",
            "**Presentation** (1 ops)");
        CorruptOnce(
            Path.Combine(temp.Path, "mcpb", "manifest.json"),
            @"\d+ tools \(\d+ operations across \d+ domains",
            "1 tools (2 operations across 3 domains");
        CorruptOnce(
            Path.Combine(temp.Path, "gh-pages", "docs", "reference", "behavioral-rules.md"),
            @"The other \d+ tools use",
            "The other 1 tools use");

        var arguments = new[]
        {
            "-RepoRoot", RepoRoot,
            "-DocsRoot", temp.Path,
            "-SkipBuild",
            "-Update",
        };
        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            arguments);

        foreach (var (relativePath, content) in expected)
        {
            Assert.Equal(content, File.ReadAllText(Path.Combine(temp.Path, relativePath)));
        }

        var headline = System.Text.RegularExpressions.Regex.Match(
            expected["README.md"],
            @"(?<tools>\d+) MCP tools with (?<operations>\d+) operations across (?<domains>\d+) domains");
        Assert.True(headline.Success);
        using (var counts = JsonDocument.Parse(File.ReadAllText(Path.Combine(temp.Path, "doc-counts.json"))))
        {
            Assert.Equal(
                int.Parse(headline.Groups["tools"].Value, System.Globalization.CultureInfo.InvariantCulture),
                counts.RootElement.GetProperty("tools").GetInt32());
            Assert.Equal(
                int.Parse(headline.Groups["operations"].Value, System.Globalization.CultureInfo.InvariantCulture),
                counts.RootElement.GetProperty("operations").GetInt32());
            Assert.Equal(
                int.Parse(headline.Groups["domains"].Value, System.Globalization.CultureInfo.InvariantCulture),
                counts.RootElement.GetProperty("domains").GetInt32());
        }

        var validateArguments = new[]
        {
            "-RepoRoot", RepoRoot,
            "-DocsRoot", temp.Path,
            "-SkipBuild",
        };
        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            validateArguments);

        CorruptOnce(
            Path.Combine(temp.Path, "README.md"),
            @"\d+ MCP tools with \d+ operations",
            "1 MCP tools with 2 operations");
        var stale = RunPowerShellRaw(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            validateArguments);
        Assert.NotEqual(0, stale.ExitCode);
        Assert.Contains("README.md", stale.Output, StringComparison.Ordinal);
        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            [.. validateArguments, "-AllowStaleAdvertisedCounts"]);

        var docCountsPath = Path.Combine(temp.Path, "doc-counts.json");
        var staleCounts = System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(docCountsPath))!.AsObject();
        staleCounts["tools"] = 1;
        File.WriteAllText(docCountsPath, staleCounts.ToJsonString());
        var staleDocCounts = RunPowerShellRaw(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            validateArguments);
        Assert.NotEqual(0, staleDocCounts.ExitCode);
        Assert.Contains("doc-counts.json", staleDocCounts.Output, StringComparison.Ordinal);
        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            [.. validateArguments, "-AllowStaleAdvertisedCounts"]);

        var incompatible = RunPowerShellRaw(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            [.. validateArguments, "-Update", "-AllowStaleAdvertisedCounts"]);
        Assert.NotEqual(0, incompatible.ExitCode);
        Assert.Contains("cannot be used together", incompatible.Output, StringComparison.Ordinal);

        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            [.. validateArguments, "-Update"]);
        var updatedDocCounts = File.ReadAllText(docCountsPath);
        RunPowerShell(
            Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"),
            [.. validateArguments, "-Update"]);
        Assert.Equal(updatedDocCounts, File.ReadAllText(docCountsPath));
        foreach (var (relativePath, content) in expected)
        {
            Assert.Equal(content, File.ReadAllText(Path.Combine(temp.Path, relativePath)));
        }
    }

    [Fact]
    public void DocumentationCountWorkflow_UpdatesCountsOnMainAndReleaseValidatesThem()
    {
        var workflow = File.ReadAllText(ReleaseWorkflow);
        var docCountsWorkflow = File.ReadAllText(Path.Combine(
            RepoRoot,
            ".github",
            "workflows",
            "doc-counts.yml"));

        Assert.Contains("prepare-release-docs:", workflow, StringComparison.Ordinal);
        Assert.Contains("./scripts/check-doc-counts.ps1 -SkipBuild", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("Update-DocumentationCounts.ps1", workflow, StringComparison.Ordinal);
        Assert.Contains("name: generated-documentation", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("skills/shared", workflow, StringComparison.Ordinal);
        Assert.Contains("skills/powerpoint-deck-design/SKILL.md", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("PopulateReferences", workflow, StringComparison.Ordinal);
        Assert.Contains("path: .", workflow, StringComparison.Ordinal);
        Assert.Contains("git add CHANGELOG.md package.json .changeset", workflow, StringComparison.Ordinal);
        Assert.Contains("git add --update", workflow, StringComparison.Ordinal);

        Assert.Contains("branches: [main]", docCountsWorkflow, StringComparison.Ordinal);
        Assert.Contains("check-doc-counts.ps1 -Update", docCountsWorkflow, StringComparison.Ordinal);
        Assert.DoesNotContain("Build-AgentSkills.ps1", docCountsWorkflow, StringComparison.Ordinal);
        Assert.Matches(
            @"git diff --cached --quiet\s+if \(\$LASTEXITCODE -eq 0\) \{",
            docCountsWorkflow);
        Assert.Contains(
            "if ($LASTEXITCODE -ne 1) {",
            docCountsWorkflow,
            StringComparison.Ordinal);
        Assert.Contains("git push origin HEAD:main", docCountsWorkflow, StringComparison.Ordinal);
        Assert.Contains(
            "check-doc-counts.ps1 -SkipBuild -AllowStaleAdvertisedCounts",
            File.ReadAllText(CiWorkflow),
            StringComparison.Ordinal);
        Assert.DoesNotContain("check-doc-counts.ps1", File.ReadAllText(PreCommitScript), StringComparison.Ordinal);
    }

    [Fact]
    public void PreCommit_SkillGenerationTestsDoNotTriggerFullMcpSuite()
    {
        var planner = Path.Combine(RepoRoot, "scripts", "Get-ValidationPlan.ps1");
        var command = $"""
            . '{planner}';
            @(
                Get-ValidationPlan -Paths 'tests/PowerPointMcp.SkillGeneration.Tests/ReleasePackagingTests.cs';
                Get-ValidationPlan -Paths 'tests/PowerPointMcp.McpServer.Tests/Integration/McpProtocolTests.cs';
                Get-ValidationPlan -Paths 'src/PowerPointMcp.Core/Slide/SlideCommands.cs';
                Get-ValidationPlan -Paths 'vscode-extension/src/extension.ts';
                Get-ValidationPlan -Paths '.github/plugins/powerpoint-mcp/mcp.json'
            ) | ConvertTo-Json -Compress
            """;
        var result = RunProcessRaw("pwsh", ["-NoProfile", "-Command", command]);

        Assert.Equal(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.Output);
        var plans = document.RootElement.EnumerateArray().ToArray();
        Assert.False(plans[0].GetProperty("PowerPoint").GetBoolean());
        Assert.True(plans[0].GetProperty("SkillTests").GetBoolean());
        Assert.True(plans[1].GetProperty("PowerPoint").GetBoolean());
        Assert.True(plans[2].GetProperty("PowerPoint").GetBoolean());
        Assert.True(plans[3].GetProperty("Extension").GetBoolean());
        Assert.True(plans[4].GetProperty("Plugins").GetBoolean());
    }

    [Fact]
    public void PowerPointFreeTestSelection_ProjectFilesRunAllGroups()
    {
        foreach (var project in new[] { "CLI", "McpServer", "Generators.Cli", "Generators.Mcp" })
        {
            AssertSelectedPowerPointFreeTests(
                [$"src/PowerPointMcp.{project}/PowerPointMcp.{project}.csproj"],
                ["CLI", "McpServer", "SkillGeneration"]);
        }
        AssertSelectedPowerPointFreeTests(
            ["tests/PowerPointMcp.CLI.Tests/PowerPointMcp.CLI.Tests.csproj"],
            ["CLI", "McpServer", "SkillGeneration"]);
    }

    [Fact]
    public void PowerPointFreeTestSelection_ChoosesAffectedGroupsAndFallsBackSafely()
    {
        AssertSelectedPowerPointFreeTests(
            ["tests/PowerPointMcp.McpServer.Tests/Integration/McpProtocolTests.cs"],
            ["McpServer"]);
        AssertSelectedPowerPointFreeTests(
            ["tests/PowerPointMcp.CLI.Tests/CommandTests.cs"],
            ["CLI"]);
        AssertSelectedPowerPointFreeTests(
            ["src/PowerPointMcp.Core/Slide/SlideCommands.cs"],
            ["CLI", "McpServer", "SkillGeneration"]);
        AssertSelectedPowerPointFreeTests(
            ["skills/powerpoint-mcp/SKILL.md"],
            ["SkillGeneration"]);
        AssertSelectedPowerPointFreeTests(
            ["docs/usage.md", "README.md", ".changeset/session-id.md"],
            []);
        AssertSelectedPowerPointFreeTests(
            ["new-root-config.bin"],
            ["CLI", "McpServer", "SkillGeneration"]);
        AssertSelectedPowerPointFreeTests(
            [],
            ["CLI", "McpServer", "SkillGeneration"]);
    }

    [Fact]
    public void CiWorkflow_SelectsPowerPointFreeTestsFromCompleteDiffOrRunsAll()
    {
        var workflow = File.ReadAllText(CiWorkflow);

        Assert.Contains("fetch-depth: 0", workflow, StringComparison.Ordinal);
        Assert.Contains("PULL_REQUEST_BASE_SHA", workflow, StringComparison.Ordinal);
        Assert.Contains("PUSH_BEFORE_SHA", workflow, StringComparison.Ordinal);
        Assert.Contains("No reliable change range is available; all PowerPoint-free test groups will run.", workflow, StringComparison.Ordinal);
        Assert.Contains("-SelectChangedPaths -ChangedPaths $changedPaths", workflow, StringComparison.Ordinal);
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
        const string runner = "Invoke-PowerPointFreeTests.ps1";

        Assert.Contains(runner, File.ReadAllText(CiWorkflow), StringComparison.Ordinal);
        Assert.Contains(runner, File.ReadAllText(PreCommitScript), StringComparison.Ordinal);
    }

    [Fact]
    public void VscodePackaging_BuildsAndPublishesBothNativeTargets()
    {
        var ci = File.ReadAllText(CiWorkflow);
        var release = File.ReadAllText(ReleaseWorkflow);
        var builder = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "Build-VscodeExtension.ps1"));
        var helpers = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"));

        Assert.Contains("Build-VscodeExtension.ps1", ci, StringComparison.Ordinal);
        Assert.Contains("Build-VscodeExtension.ps1", release, StringComparison.Ordinal);
        Assert.Contains("win32-x64", builder, StringComparison.Ordinal);
        Assert.Contains("win32-arm64", builder, StringComparison.Ordinal);
        Assert.Contains("Assert-PackageRuntimeArchitecture", builder, StringComparison.Ordinal);
        Assert.Contains("0x8664", helpers, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("0xAA64", helpers, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            "powerpoint-mcp-$env:VERSION.vsix\" --skip-duplicate",
            release,
            StringComparison.Ordinal);
        Assert.Contains(
            "powerpoint-mcp-$env:VERSION-win32-arm64.vsix\" --skip-duplicate",
            release,
            StringComparison.Ordinal);
        Assert.DoesNotContain("continue-on-error: true", release, StringComparison.Ordinal);
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
    [InlineData("for (dynamic item = source.Child; item != null; ) { break; }", 1, 1)]
    [InlineData("for (dynamic item = source.Child; item != null; ) { ComUtilities.Release(ref item); break; }", 0, 1)]
    [InlineData("for (dynamic item = source.Child; item != null; ) { break; } for (dynamic item = source.Other; item != null; ) { ComUtilities.Release(ref item); break; }", 1, 2)]
    [InlineData("using (dynamic item = source.Child) { }", 1, 1)]
    [InlineData("using dynamic item = source.Child;", 1, 1)]
    [InlineData("dynamic borrowed = (ctx).Presentation;", 0, 0)]
    [InlineData("dynamic borrowed = ((dynamic)ctx).Presentation;", 0, 0)]
    [InlineData("dynamic borrowed = context!.App;", 0, 0)]
    [InlineData("dynamic borrowed = ((dynamic)(context!)).App;", 0, 0)]
    [InlineData("dynamic item = (ctx).Presentation.Slides;", 1, 1)]
    [InlineData("dynamic item = ((dynamic)context).App.Presentations;", 1, 1)]
    [InlineData("dynamic? item = null; Retry(() => { item = source.Child; });", 1, 1)]
    [InlineData("dynamic? item = null; Retry(() => { item = source.Child; }); ComUtilities.Release(ref item);", 0, 1)]
    [InlineData("dynamic? item = null; void Acquire() { item = source.Child; } Acquire();", 1, 1)]
    [InlineData("dynamic? item = null; Retry(() => { dynamic item = null; item = source.Child; ComUtilities.Release(ref item); });", 0, 1)]
    [InlineData("dynamic? item = null; Retry((dynamic item) => { item = source.Child; });", 0, 0)]
    [InlineData("dynamic item = source.Child; other.ComUtilities.Release(ref item);", 1, 1)]
    [InlineData("dynamic item = source.Child; ComUtilities.ReleaseIfNotNull(ref item);", 1, 1)]
    [InlineData("dynamic item = source.Child; Sbroenne.PowerPointMcp.ComInterop.ComUtilities.Release(ref item);", 0, 1)]
    [InlineData("dynamic item = source.Child; global::Sbroenne.PowerPointMcp.ComInterop.ComUtilities.Release(ref item);", 0, 1)]
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

    private static readonly string[] DocumentationCountPaths =
    [
        "README.md",
        Path.Combine("docs", "POWERPOINT-NATIVE-FEATURE-AUDIT.md"),
        Path.Combine("src", "PowerPointMcp.McpServer", "README.md"),
        Path.Combine("mcpb", "README.md"),
        Path.Combine("mcpb", "manifest.json"),
        Path.Combine("gh-pages", "docs", "index.md"),
        Path.Combine("gh-pages", "docs", "installation.md"),
        Path.Combine("gh-pages", "docs", "features.md"),
        Path.Combine("gh-pages", "docs", "mcp-server.md"),
        Path.Combine("gh-pages", "docs", "reference", "behavioral-rules.md"),
        Path.Combine("gh-pages", "docs", "reference", "workflows.md"),
    ];

    private static void CorruptOnce(string path, string pattern, string replacement)
    {
        var original = File.ReadAllText(path);
        var corrupted = new System.Text.RegularExpressions.Regex(pattern)
            .Replace(original, replacement, 1);
        Assert.NotEqual(original, corrupted);
        File.WriteAllText(path, corrupted);
    }

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

    private static void AssertSelectedPowerPointFreeTests(
        string[] changedPaths,
        string[] expectedGroups)
    {
        var script = Path.Combine(RepoRoot, "scripts", "Invoke-PowerPointFreeTests.ps1");
        var quotedPaths = string.Join(
            ", ",
            changedPaths.Select(path => $"'{path.Replace("'", "''", StringComparison.Ordinal)}'"));
        var command = $"& '{script}' -SelectChangedPaths -ListOnly -ChangedPaths @({quotedPaths})";
        var result = RunProcessRaw("pwsh", ["-NoProfile", "-Command", command]);
        Assert.Equal(0, result.ExitCode);

        var output = result.Output.Trim();
        var actualGroups = output.Length == 0
            || output == "No PowerPoint-free test groups selected."
            ? []
            : output.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        Assert.Equal(expectedGroups.Order(StringComparer.Ordinal), actualGroups.Order(StringComparer.Ordinal));
    }

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
