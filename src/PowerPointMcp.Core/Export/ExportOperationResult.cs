namespace Sbroenne.PowerPointMcp.Core.Export;

/// <summary>
/// Result of a slide-export operation (export to image).
/// </summary>
/// <remarks>
/// Follows the same Success/ErrorMessage invariant as other domain results (Rule 1):
/// Success == true implies ErrorMessage is null/empty. Never set both.
/// </remarks>
public sealed class ExportOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false; null/empty when Success is true.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Full path to the exported image file. Set by single-slide export; null for
    /// multi-slide export (see <see cref="ExportedFilePaths"/>).
    /// </summary>
    public string? ExportedFilePath { get; init; }

    /// <summary>
    /// Ordered list of full paths to all exported image files. Set by multi-slide export
    /// (<see cref="IExportCommands.ExportAllSlidesToImages"/>); null for single-slide export.
    /// </summary>
    public IReadOnlyList<string>? ExportedFilePaths { get; init; }

    /// <summary>Number of slides exported. Null when the operation did not complete.</summary>
    public int? SlideCount { get; init; }
}
