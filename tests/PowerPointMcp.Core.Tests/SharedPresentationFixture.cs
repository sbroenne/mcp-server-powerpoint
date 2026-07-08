using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// xUnit class fixture that keeps ONE live PowerPoint.Application instance (and its dedicated
/// STA thread) alive for the lifetime of a whole test class, instead of every [Fact] paying its
/// own PowerPoint launch + <see cref="PresentationShutdownService"/> teardown cost. Teardown
/// legitimately includes PowerPoint's documented ~90-100s post-Quit() process-exit lingering
/// (Ripley's MCP round-trip finding, .squad/decisions.md 2026-07-01) — that cost is now paid
/// ONCE per test class (at fixture disposal) instead of once per test method.
/// </summary>
/// <remarks>
/// Isolation is preserved at the FILE level, not the process level:
/// <list type="bullet">
/// <item><see cref="CreateFreshPresentation"/> closes whatever presentation is currently open
/// (WITHOUT quitting the Application) and creates a brand-new, uniquely-named blank
/// presentation in the SAME Application instance. Call this at the start of every [Fact].</item>
/// <item><see cref="ReopenCurrentPresentation"/> closes and reopens the SAME file from disk —
/// for "did my Save() actually persist to disk" round-trip assertions — again without spinning
/// up a new PowerPoint process.</item>
/// </list>
/// Backed by <see cref="PresentationBatch.ReopenPresentation"/>, an internal-only affordance
/// (the ComInterop project grants InternalsVisibleTo to this test assembly) that is
/// deliberately NOT part of the public <see cref="IPresentationBatch"/> contract — production
/// callers still get one presentation per batch.
/// </remarks>
public sealed class SharedPresentationFixture : IDisposable
{
    private readonly List<string> _createdPaths = [];
    private readonly PresentationBatch _batch;
    private string _currentPath;

    public SharedPresentationFixture()
    {
        // Must run first, before this fixture's own batch exists — see SharedTemplateAsset's
        // remarks for why building the shared .potx template requires strict, non-overlapping
        // PowerPoint process lifetimes. Idempotent/cheap on every call after the first.
        SharedTemplateAsset.EnsureBuilt();

        _currentPath = CoreTestHelper.CreateUniqueTestFilePath();
        Batch = PresentationSession.CreateNew(_currentPath);
        _batch = (PresentationBatch)Batch;
        _createdPaths.Add(_currentPath);
    }

    /// <summary>The single, shared live batch/session for the whole test class.</summary>
    public IPresentationBatch Batch { get; }

    /// <summary>
    /// Closes whatever presentation is currently open (without quitting the Application) and
    /// creates a brand-new, uniquely-named blank presentation, returning its full path. Call
    /// this at the start of every [Fact] that needs a clean, isolated presentation.
    /// </summary>
    public string CreateFreshPresentation(string prefix = "pptmcp-test")
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath(prefix);
        _batch.ReopenPresentation(path, createNewFile: true);
        _currentPath = path;
        _createdPaths.Add(path);
        return path;
    }

    /// <summary>
    /// Closes and reopens the presentation at its CURRENT path (the one from the last
    /// <see cref="CreateFreshPresentation"/> call) from disk, for asserting that a prior
    /// Save() genuinely persisted changes — without launching a new PowerPoint process.
    /// </summary>
    public void ReopenCurrentPresentation()
    {
        _batch.ReopenPresentation(_currentPath, createNewFile: false);
    }

    public void Dispose()
    {
        Batch.Dispose();

        foreach (string path in _createdPaths)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { /* best-effort cleanup; not test-critical */ }
        }
    }
}
