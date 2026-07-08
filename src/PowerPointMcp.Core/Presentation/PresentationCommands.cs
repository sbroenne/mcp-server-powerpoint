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

    private static readonly string[] AcceptedTemplateExtensions = [".potx", ".potm", ".pot", ".pptx", ".pptm", ".ppt"];

    /// <inheritdoc/>
    public PresentationOperationResult ApplyTemplate(IPresentationBatch batch, string templatePath)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (string.IsNullOrWhiteSpace(templatePath))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = "A template path is required."
            };
        }

        string fullTemplatePath = Path.GetFullPath(templatePath);
        string extension = Path.GetExtension(fullTemplatePath);

        // Rule 1b: a missing file or unsupported extension is expected/graceful bad input —
        // validate up front and fail without ever calling into COM. Unexpected COM failures
        // (e.g. a corrupt template PowerPoint can't parse) are NOT caught here and propagate
        // from ApplyTemplate below.
        if (!File.Exists(fullTemplatePath))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = $"Template file not found: '{fullTemplatePath}'."
            };
        }

        if (!AcceptedTemplateExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = $"'{extension}' is not a supported template extension. Expected one of: {string.Join(", ", AcceptedTemplateExtensions)}."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            ctx.Presentation.ApplyTemplate(fullTemplatePath);

            string? themeName = ctx.Presentation.Designs.Count > 0
                ? ctx.Presentation.Designs[1].Name
                : null;

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath,
                ThemeName = themeName
            };
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult GetThemeName(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            string? themeName = ctx.Presentation.Designs.Count > 0
                ? ctx.Presentation.Designs[1].Name
                : null;

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath,
                ThemeName = themeName
            };
        });
    }
}
