using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Static code-quality guards for the PIA migration: dynamic-keyword ratchet,
/// reflection-activation ban, and raw-pp*-constant ban for non-Image Core/ComInterop.
/// </summary>
/// <remarks>
/// <para><b>Dynamic-keyword ratchet</b> (<see cref="NonImageCoreDomains_DynamicKeywordCount_DoesNotExceedAllowlist"/>):
/// asserts that the count of <c>dynamic</c> keyword usages in each non-Image Core source file
/// does not EXCEED the actual count captured after the PIA migration (2026-07-11). Counts are
/// set to the real post-migration values, making this a true one-way ratchet:
/// decreasing counts keep it green (progress), increasing counts fail it (regression).
/// </para>
///
/// <para><b>Reflection-activation ban</b> (<see cref="NonImageCoreAndComInterop_ReflectionActivationPatterns_AreAbsent"/>):
/// scans non-Image Core <em>and</em> ComInterop (excluding <c>obj/</c>) for known COM-bypassing
/// reflection patterns: <c>Type.GetTypeFromProgID</c>, <c>Activator.CreateInstance</c>,
/// <c>.InvokeMember(</c>. These patterns circumvent the embedded PowerPoint PIA and must not
/// appear in production code.</para>
///
/// <para><b>Raw pp* constant ban</b> (<see cref="NonImageCore_RawPpEnumConstantDefinitions_AreAbsent"/>):
/// scans non-Image Core for assignments of the form <c>pp[A-Z]\w+ = \d+</c> — a raw integer
/// constant definition for a pp*-named identifier that should instead use a typed
/// <c>PowerPoint.Pp*</c> PIA value. Limitation: inline literal comparisons (e.g.
/// <c>shape.Type == 2</c>) are not caught by this guard — those require code review. Justified
/// <c>mso*</c> integer constants (e.g. <c>const int MsoShapePlaceholder = 14</c>) are not
/// matched and remain permitted.</para>
///
/// <para><b>Rules:</b></para>
/// <list type="bullet">
///   <item>A file's dynamic count <em>may decrease</em> (migration progress). Green.</item>
///   <item>A file's dynamic count <em>may not increase</em> without updating the allowlist. Red.</item>
///   <item>Image domain is <em>hard-excluded</em> from all guards — a separate task owns it.</item>
///   <item>Reflection-activation and raw-pp* guards have zero tolerance — no allowlist, fail on
///   first match.</item>
/// </list>
///
/// <para><b>How to update the dynamic allowlist legitimately:</b> (1) verify the new usage
/// cannot use typed PIA; (2) add an inline comment in the source file explaining why;
/// (3) increment the matching max in <see cref="DynamicAllowlist"/> and reference the reason.</para>
///
/// <para><b>Note on accuracy:</b> the regex <c>\bdynamic\s</c> also matches occurrences inside
/// XML doc comments — these are rare in this codebase and remain within the allowlist budget.
/// The guard is intentionally approximate (preventing drift, not perfect static analysis).</para>
/// </remarks>
[Trait("Category", "CodeQuality")]
public class PiaMigrationGuardTests
{
    // Per-file maximum allowed count of `dynamic ` keyword occurrences (word boundary + whitespace).
    // Counts are set to the ACTUAL post-migration values (2026-07-11) so this is a true one-way
    // ratchet: a file's count may decrease (migration progress → stays green) but may not increase
    // (regression → fails).
    //
    // Justified residual dynamic per file:
    //   Animation  — Effect.Exit and SlideShowTransition.AdvanceOnClick / .AdvanceOnTime are
    //                MsoTriState (Office.Core / office.dll, not referenced); the typed setter
    //                requires dynamic dispatch. MsoAnimEffect, MsoAnimTriggerType, PpEntryEffect,
    //                and MsoAnimateByLevel are all exposed by the embedded PowerPoint PIA and have
    //                already been migrated to typed form — they are NOT the reason for residual
    //                dynamic here.
    //   Chart      — Shape.HasChart is MsoTriState (Office.Core); SeriesCollection / XValues /
    //                NewSeries are Microsoft.Office.Interop.Excel types (Excel.dll not referenced);
    //                chart-shape access via dynamic to avoid that reference. Residual count
    //                reflects the post-migration state where only these office.dll-dependent paths
    //                remain dynamic.
    //   Master     — ctx.Presentation.SlideMaster hangs indefinitely via typed NoPIA interface
    //                (documented quirk identical to CoreTestHelper.CreateTemplateFile). Stays dynamic.
    //   Notes      — NotesPage / shape iteration accessed dynamically; msoPlaceholder (= 14) is an
    //                MsoShapeType raw int constant (no typed equivalent in scope without office.dll);
    //                PpPlaceholderType is already typed via PowerPoint PIA.
    //   Presentation — BuiltInDocumentProperties / CustomDocumentProperties return
    //                  Microsoft.Office.Core.DocumentProperty (office.dll). Cannot use typed PIA.
    //   Shape      — MsoAutoShapeType / MsoConnectorType / MsoTextOrientation / MsoShadowStyle /
    //                MsoReflectionType etc. are all in office.dll. Line.Visible and Shadow.Visible
    //                setters are MsoTriState (office.dll). Residual dynamic is the post-migration
    //                minimum for these office.dll-dependent properties.
    //   Slide      — FollowMasterBackground is MsoTriState (office.dll); SectionProperties /
    //                FillFormat.TwoColorGradient paths also require dynamic for office.dll types.
    //   SmartArt   — Office.SmartArtLayouts / SmartArtNode / SmartArtNodes are in office.dll;
    //                Shape.HasSmartArt is MsoTriState (office.dll). Entire SmartArt surface stays
    //                dynamic by design.
    //   Table      — Cell border Visible setter is MsoTriState (office.dll); table/cell access
    //                via dynamic for the border line-style properties with no typed PIA path.
    //   TextFrame  — Font.Bold / .Italic / .Underline and ParagraphFormat.Bullet signatures
    //                reference Office.Core types from office.dll. Minimum residual count.
    private static readonly Dictionary<string, int> DynamicAllowlist = new(StringComparer.OrdinalIgnoreCase)
    {
        // domain path (relative to src/PowerPointMcp.Core)         max allowed   reason
        ["Animation\\AnimationCommands.cs"]       =  4,  // Effect.Exit + Transition.AdvanceOnClick/AdvanceOnTime are MsoTriState (office.dll)
        ["Chart\\ChartCommands.cs"]               = 21,  // Shape.HasChart (MsoTriState) + Excel.SeriesCollection/XValues (Excel.dll)
        ["Master\\MasterCommands.cs"]             =  5,  // SlideMaster NoPIA hang + Font.Bold MsoTriState (office.dll)
        ["Notes\\NotesCommands.cs"]               =  3,  // NotesPage / shape iteration — office.dll types
        ["Presentation\\PresentationCommands.cs"] =  5,  // DocumentProperties (office.dll)
        ["Shape\\ShapeCommands.cs"]               = 26,  // MsoAutoShapeType / MsoShadowStyle / Line+Shadow Visible (office.dll)
        ["Slide\\SlideCommands.cs"]               =  5,  // FollowMasterBackground MsoTriState + FillFormat/SectionProperties (office.dll)
        ["SmartArt\\SmartArtCommands.cs"]         = 20,  // SmartArtLayouts / HasSmartArt MsoTriState (office.dll)
        ["Table\\TableCommands.cs"]               =  4,  // Border Visible MsoTriState + cell/border access (office.dll)
        ["TextFrame\\TextFrameCommands.cs"]       =  7,  // Font tri-state + ParagraphFormat.Bullet (office.dll)
    };

    /// <summary>
    /// For every non-Image Core .cs file in <see cref="DynamicAllowlist"/>, asserts that the
    /// actual count of <c>dynamic</c> keyword usages does not exceed the recorded maximum.
    /// Migration progress (decreasing counts) passes; new additions (increasing counts) fail.
    /// </summary>
    [Fact]
    public void NonImageCoreDomains_DynamicKeywordCount_DoesNotExceedAllowlist()
    {
        string repoRoot = FindRepoRoot();
        string coreSourceRoot = Path.Combine(repoRoot, "src", "PowerPointMcp.Core");

        var violations = new List<string>();
        var missing = new List<string>();

        foreach (var (relPath, maxAllowed) in DynamicAllowlist)
        {
            string fullPath = Path.Combine(coreSourceRoot, relPath);
            if (!File.Exists(fullPath))
            {
                missing.Add($"  {relPath} — file not found (renamed/deleted?). Update DynamicAllowlist.");
                continue;
            }

            string content = File.ReadAllText(fullPath);
            int actual = CountDynamicKeyword(content);

            if (actual > maxAllowed)
            {
                violations.Add(
                    $"  {relPath}: found {actual}, max {maxAllowed}. " +
                    "Add a code comment explaining why the new 'dynamic' is justified, then increment its max in DynamicAllowlist.");
            }
        }

        var allErrors = violations.Concat(missing).ToList();
        Assert.True(
            allErrors.Count == 0,
            "PIA migration guard failed.\n\n" +
            string.Join("\n", allErrors) + "\n\n" +
            "This test prevents NEW unjustified late-binding during the PIA migration. " +
            "To legitimately add a new 'dynamic' usage: (1) verify it cannot use typed PIA, " +
            "(2) add an inline comment in the source file explaining why, " +
            "(3) increment the matching max in DynamicAllowlist in PiaMigrationGuardTests.cs.");
    }

    /// <summary>
    /// Verifies that the Image domain is not accidentally included in the allowlist.
    /// Image is owned by a separate migration task and must never appear here.
    /// </summary>
    [Fact]
    public void ImageDomain_IsNotInAllowlist()
    {
        foreach (string key in DynamicAllowlist.Keys)
        {
            Assert.False(
                key.Contains("Image", StringComparison.OrdinalIgnoreCase),
                $"Image domain path '{key}' must not appear in the PIA migration guard allowlist. " +
                "Image is owned by a separate task (hard exclusion).");
        }
    }

    /// <summary>
    /// Scans non-Image Core <em>and</em> ComInterop (excluding generated <c>obj/</c> output)
    /// for known reflection-based COM activation/dispatch patterns that bypass the embedded
    /// PowerPoint PIA:
    /// <list type="bullet">
    ///   <item><c>Type.GetTypeFromProgID(</c> — late-bound ProgID activation</item>
    ///   <item><c>Activator.CreateInstance(</c> — generic reflection instantiation</item>
    ///   <item><c>.InvokeMember(</c> — late-bound member dispatch</item>
    /// </list>
    /// Zero occurrences are allowed. PowerPoint should always be started via the typed
    /// <c>new PowerPoint.Application()</c> PIA constructor (as in <c>PresentationBatch.cs</c>),
    /// never via reflection activation.
    /// </summary>
    [Fact]
    public void NonImageCoreAndComInterop_ReflectionActivationPatterns_AreAbsent()
    {
        string repoRoot = FindRepoRoot();
        string coreRoot      = Path.Combine(repoRoot, "src", "PowerPointMcp.Core");
        string comInteropRoot = Path.Combine(repoRoot, "src", "PowerPointMcp.ComInterop");

        var coreFiles = Directory.GetFiles(coreRoot, "*.cs", SearchOption.AllDirectories)
            .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}Image{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
                     && !f.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase));

        var comInteropFiles = Directory.GetFiles(comInteropRoot, "*.cs", SearchOption.AllDirectories)
            .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase));

        var allFiles = coreFiles.Concat(comInteropFiles);

        // Patterns that represent COM activation/dispatch via reflection — all are forbidden.
        var forbiddenPatterns = new (string Label, Regex Pattern)[]
        {
            ("Type.GetTypeFromProgID",    new Regex(@"\bType\.GetTypeFromProgID\s*\(",    RegexOptions.Compiled)),
            ("Activator.CreateInstance",  new Regex(@"\bActivator\.CreateInstance\s*\(",  RegexOptions.Compiled)),
            (".InvokeMember",             new Regex(@"\.InvokeMember\s*\(",               RegexOptions.Compiled)),
        };

        var violations = new List<string>();

        foreach (string filePath in allFiles)
        {
            string content = File.ReadAllText(filePath);
            foreach (var (label, pattern) in forbiddenPatterns)
            {
                if (pattern.IsMatch(content))
                {
                    string rel = GetRelativePath(repoRoot, filePath);
                    int count  = pattern.Count(content);
                    violations.Add($"  {rel}: found {count}× '{label}'");
                }
            }
        }

        Assert.True(
            violations.Count == 0,
            "Reflection-based COM activation/dispatch found in Core or ComInterop.\n\n" +
            string.Join("\n", violations) + "\n\n" +
            "Use typed PowerPoint PIA ('new PowerPoint.Application()') instead of " +
            "Type.GetTypeFromProgID / Activator.CreateInstance / InvokeMember.");
    }

    /// <summary>
    /// Scans non-Image Core for raw integer constant <em>definitions</em> of pp*-named
    /// identifiers (e.g. <c>const int ppPlaceholderBody = 2;</c>), which should instead use
    /// typed <c>PowerPoint.Pp*</c> PIA values that are available in the embedded assembly.
    /// </summary>
    /// <remarks>
    /// Pattern matched: <c>\bpp[A-Z]\w+\s*=\s*\d+</c>
    /// (a pp*-prefixed identifier directly assigned a raw integer literal).
    ///
    /// <b>Not matched (and not rejected):</b>
    /// <list type="bullet">
    ///   <item><c>mso*</c>-named integer constants — e.g. <c>const int MsoShapePlaceholder = 14</c>
    ///   — these refer to Office.Core types not exposed in the embedded PowerPoint PIA and are
    ///   a documented justified exception (Notes, Master).</item>
    ///   <item>Typed PIA usages — e.g. <c>= PowerPoint.PpEntryEffect.ppEffectFade</c> — the
    ///   right-hand side is not a bare integer.</item>
    ///   <item>String-literal keys — e.g. <c>["ppEffectFade"]</c> in dictionary initializers —
    ///   the word boundary + `\s*=\s*\d+` suffix prevents a match.</item>
    /// </list>
    ///
    /// <b>Known limitation:</b> inline literal <em>comparisons</em> such as
    /// <c>shape.Type == 2</c> (where 2 should be <c>PowerPoint.PpPlaceholderType.ppPlaceholderBody</c>)
    /// are not reliably detectable without parsing context and are therefore not guarded here.
    /// Those patterns must be caught via code review.
    /// </remarks>
    [Fact]
    public void NonImageCore_RawPpEnumConstantDefinitions_AreAbsent()
    {
        string repoRoot  = FindRepoRoot();
        string coreRoot  = Path.Combine(repoRoot, "src", "PowerPointMcp.Core");

        var coreFiles = Directory.GetFiles(coreRoot, "*.cs", SearchOption.AllDirectories)
            .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}Image{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
                     && !f.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase));

        // Matches: bare pp[A-Z]-prefixed identifier immediately followed by `= <digits>`
        // e.g. `ppPlaceholderBody = 2` — should use PowerPoint.PpPlaceholderType.ppPlaceholderBody
        var rawPpPattern = new Regex(@"\bpp[A-Z]\w+\s*=\s*\d+", RegexOptions.Compiled);

        var violations = new List<string>();

        foreach (string filePath in coreFiles)
        {
            string content = File.ReadAllText(filePath);
            var matches = rawPpPattern.Matches(content);
            if (matches.Count > 0)
            {
                string rel = GetRelativePath(repoRoot, filePath);
                foreach (Match m in matches)
                {
                    violations.Add($"  {rel}: raw pp* constant '{m.Value.Trim()}' — use typed PowerPoint.Pp* PIA value instead");
                }
            }
        }

        Assert.True(
            violations.Count == 0,
            "Raw pp* integer constant definitions found in non-Image Core. " +
            "These should use typed PowerPoint PIA values (e.g. PowerPoint.PpPlaceholderType.ppPlaceholderBody).\n\n" +
            string.Join("\n", violations) + "\n\n" +
            "Justified mso* integer constants (e.g. MsoShapePlaceholder = 14) remain permitted " +
            "and are not matched by this guard. See test remarks for the known limitation on " +
            "inline literal comparisons.");
    }

    // ---------------------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Counts occurrences of the <c>dynamic</c> keyword (word boundary followed by whitespace)
    /// in <paramref name="source"/>. Counts inside XML doc comments are included; this is
    /// intentional and acceptable given the low frequency in this codebase.
    /// </summary>
    private static int CountDynamicKeyword(string source)
        => Regex.Count(source, @"\bdynamic\s");

    /// <summary>
    /// Returns a path relative to <paramref name="root"/>, for readable assertion messages.
    /// Falls back to the full path when no common prefix is found.
    /// </summary>
    private static string GetRelativePath(string root, string fullPath)
    {
        if (fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            string rel = fullPath.Substring(root.Length);
            return rel.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        return fullPath;
    }

    /// <summary>
    /// Walks up the directory tree from the executing test assembly until it finds a directory
    /// that contains a <c>.slnx</c> or <c>.sln</c> file, which identifies the repository root.
    /// </summary>
    private static string FindRepoRoot()
    {
        string? dir = Path.GetDirectoryName(typeof(PiaMigrationGuardTests).Assembly.Location);

        while (dir is not null)
        {
            bool hasSlnx = Directory.GetFiles(dir, "*.slnx", SearchOption.TopDirectoryOnly).Length > 0;
            bool hasSln  = Directory.GetFiles(dir, "*.sln",  SearchOption.TopDirectoryOnly).Length > 0;

            if (hasSlnx || hasSln)
            {
                return dir;
            }

            dir = Path.GetDirectoryName(dir);
        }

        throw new InvalidOperationException(
            "Could not locate the repository root by searching for a .slnx/.sln file. " +
            $"Test assembly location: {typeof(PiaMigrationGuardTests).Assembly.Location}");
    }
}
