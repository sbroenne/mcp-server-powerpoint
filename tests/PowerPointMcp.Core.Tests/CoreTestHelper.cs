namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Shared helpers for integration tests that need a real, unique .pptx file on disk.
/// Mirrors mcp-server-excel's CoreTestHelper pattern: every test gets its own file so
/// tests never share PowerPoint sessions or contend for file locks.
/// </summary>
public static class CoreTestHelper
{
    /// <summary>
    /// Returns a full path to a new, unique .pptx file name in the OS temp directory.
    /// Does not create the file — callers create it via PresentationSession.CreateNew.
    /// </summary>
    public static string CreateUniqueTestFilePath(string prefix = "pptmcp-test")
    {
        string dir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, $"{prefix}-{Guid.NewGuid():N}.pptx");
    }

    /// <summary>
    /// Writes a minimal valid 1x1-pixel PNG to a unique temp file and returns its path.
    /// Used by image-domain tests that need a real, readable image file on disk.
    /// </summary>
    public static string CreateUniqueTestImageFile()
    {
        string dir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, $"pptmcp-test-image-{Guid.NewGuid():N}.png");

        // Minimal valid 1x1 transparent PNG.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
        File.WriteAllBytes(path, Convert.FromBase64String(base64Png));
        return path;
    }

    /// <summary>
    /// Returns a full path to a new, unique .potx file name in the OS temp directory.
    /// Does not create the file — callers build a real template via PowerPoint COM
    /// (<c>Presentation.SaveAs(path, PpSaveAsFileType.ppSaveAsTemplate)</c>) and save it there,
    /// since there's no lightweight way to hand-author a valid .potx without driving PowerPoint.
    /// </summary>
    public static string CreateUniqueTestTemplateFilePath(string prefix = "pptmcp-test-template")
    {
        string dir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, $"{prefix}-{Guid.NewGuid():N}.potx");
    }

    /// <summary>
    /// Builds a real, minimal .potx template file via a dedicated, short-lived PowerPoint COM
    /// session: creates a blank presentation, renames its design/theme so tests can assert on a
    /// known, distinctive name after <c>ApplyTemplate</c>, then saves it as
    /// <c>PpSaveAsFileType.ppSaveAsTemplate</c>. There is no committed binary .potx test asset in
    /// this repo (no precedent for shipping binary fixtures), so this is generated on demand.
    /// Callers should cache the returned path (e.g. via a lazily-initialized static field) rather
    /// than calling this once per [Fact] — it launches its own separate PowerPoint process.
    /// </summary>
    public static string CreateTemplateFile(string designName = "PptMcpTestTemplate")
    {
        string path = CreateUniqueTestTemplateFilePath();

        using var batch = Sbroenne.PowerPointMcp.ComInterop.Session.PresentationSession.CreateNew(
            CreateUniqueTestFilePath("pptmcp-template-source"));

        batch.Execute((ctx, ct) =>
        {
            // NOTE: discovered via integration test — accessing Presentation.SlideMaster through
            // the strongly-typed, embedded (NoPIA) interop interface hangs indefinitely (the COM
            // call never returns, no dialog, process stays responsive). This matches the same
            // class of NoPIA/late-binding quirk already documented on Shapes.Index and
            // Presentation.SaveAs elsewhere in this codebase. Fix: access SlideMaster/Design/Name
            // entirely via dynamic (IDispatch) dispatch instead of the typed PIA interface.
            dynamic dynPresentation = ctx.Presentation;
            dynPresentation.SlideMaster.Design.Name = designName;
            // Use dynamic dispatch for SaveAs: the interop's optional-parameter default values
            // reference Microsoft.Office.Core.MsoTriState, which this project doesn't reference
            // (see PresentationBatch.CreateNew's identical pattern) — dynamic dispatch skips
            // static overload/default-value resolution entirely.
            dynPresentation.SaveAs(path, Sbroenne.PowerPointMcp.ComInterop.ComInteropConstants.PpSaveAsOpenXmlTemplate);
            return 0;
        });

        return path;
    }
}

/// <summary>
/// Builds the shared .potx test template exactly once, at test-assembly load time — strictly
/// BEFORE any <see cref="SharedPresentationFixture"/> (or any other test-owned
/// <c>PresentationBatch</c>) can exist.
/// </summary>
/// <remarks>
/// <para><b>Discovered via real integration testing, not assumed:</b> creating a SECOND
/// <c>PresentationBatch</c> (via <see cref="CoreTestHelper.CreateTemplateFile"/>) while another
/// one (e.g. <see cref="SharedPresentationFixture"/>'s) is already alive in the same test
/// process does NOT reliably spawn a second, independent <c>POWERPNT.exe</c>. Empirically
/// verified (via live process enumeration during a real test run): only ONE PowerPoint process
/// existed for the whole duration, meaning the "second" batch's COM activation silently reused
/// the already-running instance. When that "second" batch's <c>Dispose()</c> then calls
/// <c>Quit()</c>/force-kill as part of its normal, documented ~90-150s shutdown teardown, it
/// kills the ONE shared process out from under the first (still in-use) batch, which then fails
/// with "PowerPoint process is no longer running" on its very next COM call.</para>
/// <para>
/// The fix is to guarantee strict, non-overlapping PowerPoint process lifetimes: build the
/// template asset via its own fully-isolated <c>PresentationBatch</c> (create → save → fully
/// dispose, including PowerPoint's documented lingering post-Quit() wait) BEFORE any other test
/// fixture's batch is constructed.
/// </para>
/// <para><b>Update — a THIRD bug found and fixed:</b> a <c>[ModuleInitializer]</c> method was
/// originally used here to guarantee the "runs before anything else" ordering, since it is the
/// only point guaranteed to run before ANY type in the assembly is touched. However, real
/// integration testing showed the exact same COM call sequence (<c>SlideMaster.Design.Name</c>
/// set, then <c>SaveAs(..., ppSaveAsOpenXmlTemplate)</c>) that reliably completes in ~100-150ms
/// when invoked from a normal, already-running <c>[Fact]</c> or fixture constructor,
/// deterministically hits <c>PresentationBatch</c>'s 120s operation timeout when invoked from a
/// <c>[ModuleInitializer]</c> — and retrying inside the SAME <c>[ModuleInitializer]</c> call does
/// NOT help (both attempts hang identically), which rules out "COM just needs to warm up
/// once" as the explanation. The distinguishing factor is the CLR module-initialization
/// context itself (running during/immediately after test-assembly load, likely still under the
/// test host's own assembly-loading/reflection-discovery machinery), not merely "how early" the
/// call happens in wall-clock terms. The fix: stop using <c>[ModuleInitializer]</c> entirely.
/// Instead, <see cref="EnsureBuilt"/> lazily builds the template on first actual use, from a
/// normal (mature) execution context — a test fixture constructor — which empirically does NOT
/// hang. The "must run before any other <c>PresentationBatch</c>" ordering guarantee is now
/// provided by having <see cref="SharedPresentationFixture"/>'s constructor call
/// <see cref="EnsureBuilt"/> as its very first statement, before creating its own batch, instead
/// of relying on assembly-load-time execution.
/// </para>
/// </remarks>
internal static class SharedTemplateAsset
{
    private static readonly Lock s_lock = new();
    private static bool s_built;
    private static string? s_templatePath;
    private static Exception? s_buildError;

    /// <summary>The distinctive design/theme name baked into the shared template.</summary>
    public const string DesignName = "PptMcpTestTemplate";

    /// <summary>
    /// Builds the shared .potx template exactly once (idempotent, thread-safe), in complete
    /// isolation from every other PowerPoint COM session in this test run. Must be called
    /// BEFORE any other test fixture creates its own <c>PresentationBatch</c> — see
    /// <see cref="SharedPresentationFixture"/>'s constructor, which calls this first.
    /// </summary>
    public static void EnsureBuilt()
    {
        lock (s_lock)
        {
            if (s_built) return;

            try
            {
                s_templatePath = CoreTestHelper.CreateTemplateFile(DesignName);
                s_buildError = null;
            }
            catch (Exception ex)
            {
                s_buildError = ex;
            }
            finally
            {
                s_built = true;
            }
        }
    }

    /// <summary>
    /// Path to the shared, already-built .potx template. <see cref="EnsureBuilt"/> must have
    /// been called first (it is, from <see cref="SharedPresentationFixture"/>'s constructor).
    /// </summary>
    public static string TemplatePath => s_templatePath
        ?? throw new InvalidOperationException(
            "The shared .potx test template failed to build at test-assembly load time.",
            s_buildError);
}
