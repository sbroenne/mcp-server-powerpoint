using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <inheritdoc cref="IPresentationCommands"/>
public sealed class PresentationCommands : IPresentationCommands
{
    /// <inheritdoc/>
    public PresentationOperationResult Create(string filePath, bool isMacroEnabled = false)
    {
        // Let exceptions propagate — no try/catch suppression here (Rule 1b in mcp-server-excel's
        // instructions applies identically to this port: batch construction/Save failures should
        // surface as real exceptions to the caller layer, which is responsible for translating them
        // into CLI/MCP error results).
        using var batch = PresentationSession.CreateNew(filePath, show: false);
        batch.Save();

        return new PresentationOperationResult
        {
            Success = true,
            PresentationPath = batch.PresentationPath
        };
    }

    /// <inheritdoc/>
    public PresentationOperationResult Open(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = "A file path is required."
            };
        }

        string fullPath = Path.GetFullPath(filePath);

        // Rule 1b: a missing file is expected/graceful bad input — fail without ever starting
        // PowerPoint. Genuinely unexpected COM failures (corrupt file, PowerPoint not
        // installed, etc.) are NOT caught here and propagate from BeginBatch below.
        if (!File.Exists(fullPath))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = $"Presentation file not found: '{fullPath}'."
            };
        }

        // Open (and immediately close) a real batch to prove PowerPoint can actually open this
        // file — mirrors Create()'s create+save+close pattern. Callers that want to keep
        // editing must call PresentationSession.BeginBatch themselves and hold onto the batch.
        using var batch = PresentationSession.BeginBatch(fullPath);

        return new PresentationOperationResult
        {
            Success = true,
            PresentationPath = batch.PresentationPath
        };
    }

    /// <inheritdoc/>
    public PresentationOperationResult Save(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        batch.Save();

        return new PresentationOperationResult
        {
            Success = true,
            PresentationPath = batch.PresentationPath
        };
    }
}
