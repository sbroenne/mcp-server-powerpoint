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
}
