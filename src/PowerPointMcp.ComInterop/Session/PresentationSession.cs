using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Main entry point for PowerPoint COM interop operations using the batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class PresentationSession
{
    /// <summary>
    /// Begins a batch of PowerPoint operations against an existing presentation file.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IPresentationBatch BeginBatch(string filePath)
        => BeginBatch(filePath, show: false, operationTimeout: null);

    /// <summary>
    /// Begins a batch of PowerPoint operations against an existing presentation file, with
    /// optional UI visibility and a custom operation timeout.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IPresentationBatch BeginBatch(string filePath, bool show, TimeSpan? operationTimeout)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("A file path is required", nameof(filePath));

        string fullPath = Path.GetFullPath(filePath);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".pptx" or ".pptm" or ".ppt"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only PowerPoint files (.pptx, .pptm, .ppt) are supported.");
        }

        return new PresentationBatch(fullPath, createNewFile: false, show: show, operationTimeout: operationTimeout);
    }

    /// <summary>
    /// Creates a new PowerPoint presentation at the specified path and returns a batch with it
    /// open, ready for further operations (e.g. adding slides) before an explicit Save().
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IPresentationBatch CreateNew(string filePath, bool show = false, TimeSpan? operationTimeout = null)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("A file path is required", nameof(filePath));

        string fullPath = Path.GetFullPath(filePath);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".pptx" or ".pptm"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. New presentations must be .pptx or .pptm.");
        }

        return new PresentationBatch(fullPath, createNewFile: true, show: show, operationTimeout: operationTimeout);
    }
}
