function Get-ValidationPlan {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Paths = @())

    $plan = [ordered]@{
        Build = $false
        PowerPoint = $false
        SourceChecks = $false
        HookTests = $false
        Cli = $false
        Mcp = $false
        Extension = $false
        Mcpb = $false
        Skills = $false
        SkillTests = $false
        Plugins = $false
        Reasons = [Collections.Generic.List[string]]::new()
    }

    foreach ($original in $Paths) {
        $path = $original.Replace('\', '/')
        $kind = switch -Regex ($path) {
            '^(Directory\.Build\..*|Directory\.Packages\.props|global\.json|NuGet\.Config|Sbroenne\.PowerPointMcp\.slnx)$' { 'runtime'; break }
            '^src/PowerPointMcp\.(Core|ComInterop|Service|Generators[^/]*)/' { 'runtime'; break }
            '^src/PowerPointMcp\.CLI/' { 'cli'; break }
            '^src/PowerPointMcp\.McpServer/' { 'mcp'; break }
            '^src/PowerPointMcp\.Build\.Tasks/|^skills/' { 'skills'; break }
            '^tests/PowerPointMcp\.(Core|ComInterop|McpServer)\.Tests/' { 'runtime'; break }
            '^tests/' { 'tests'; break }
            '^vscode-extension/' { 'extension'; break }
            '^mcpb/' { 'mcpb'; break }
            '^npm-packages/pptcli' { 'cli-package'; break }
            '^npm-packages/mcp-server-powerpoint' { 'mcp-package'; break }
            '^npm-packages/shared/' { 'npm-packages'; break }
            '^\.github/plugins/|^\.github/workflows/publish-plugins\.yml$' { 'plugins'; break }
            '^scripts/Build-AgentSkills\.ps1$' { 'skills'; break }
            '^scripts/(Build-Plugins|Sync-PublishedPluginRepo)\.ps1$' { 'plugins'; break }
            '^scripts/(Build-NpmPackages|Test-NpmPackages|Build-VscodeExtension|PackageHelpers)\.ps1$|^\.github/workflows/release\.yml$' { 'packages'; break }
            '^scripts/(pre-commit|Get-ValidationPlan|Invoke-PowerPointFreeTests|check-)' { 'tests'; break }
            '^\.github/workflows/ci\.yml$' { 'pipeline'; break }
            '^scripts/(Build-Changelog|Update-(ReleaseVersion|McpRegistry)Metadata)\.ps1$' { 'packages'; break }
            '^docs/|^gh-pages/|^videos/|^infrastructure/|^specs/|^\.changeset/|^\.github/|\.md$|^doc-counts\.json$|^\.(gitignore|gitattributes)$' { 'documentation'; break }
            '^(package(-lock)?\.json|\.npmrc)$' { 'packages'; break }
            default { 'unknown' }
        }

        $plan.Reasons.Add("$path -> $kind")
        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown')) {
            $plan.Build = $true
            $plan.PowerPoint = $true
            $plan.SourceChecks = $true
        }
        if ($kind -in @('runtime', 'cli', 'cli-package', 'npm-packages', 'packages', 'pipeline', 'unknown')) { $plan.Cli = $true }
        if ($kind -in @('runtime', 'mcp', 'mcp-package', 'npm-packages', 'packages', 'pipeline', 'unknown')) { $plan.Mcp = $true }
        if ($kind -in @('runtime', 'cli', 'mcp', 'skills', 'packages', 'pipeline', 'unknown')) { $plan.Skills = $true }
        if ($kind -in @('extension', 'skills', 'packages', 'pipeline', 'runtime', 'mcp', 'unknown')) { $plan.Extension = $true }
        if ($kind -in @('mcpb', 'packages', 'pipeline', 'runtime', 'mcp', 'unknown')) { $plan.Mcpb = $true }
        if ($kind -in @('plugins', 'skills', 'packages', 'pipeline', 'runtime', 'cli', 'mcp', 'unknown')) { $plan.Plugins = $true }
        if ($kind -in @('build', 'tests', 'extension', 'mcpb', 'skills', 'plugins', 'pipeline', 'packages')) { $plan.Build = $true }
        if ($kind -in @('tests', 'pipeline')) { $plan.HookTests = $true }
        if ($kind -eq 'skills' -or $path -match '^tests/PowerPointMcp\.SkillGeneration\.Tests/') { $plan.SkillTests = $true }
    }

    [pscustomobject]$plan
}
