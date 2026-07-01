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
