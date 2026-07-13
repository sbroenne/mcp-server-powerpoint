using Sbroenne.PowerPointMcp.ComInterop;
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

    /// <summary>
    /// The built-in document properties this domain supports writing/reading, matching
    /// <c>Presentation.BuiltInDocumentProperties</c>'s name-indexed entries (verified live via
    /// COM spike). Read-only/statistical built-ins (word count, slide count, etc.) are
    /// intentionally out of scope — those are already exposed by other domains (e.g. Slide).
    /// </summary>
    private static readonly string[] SupportedBuiltInProperties =
        ["Title", "Subject", "Author", "Keywords", "Comments", "Category", "Manager", "Company"];

    /// <inheritdoc/>
    public PresentationOperationResult SetDocumentProperty(IPresentationBatch batch, string propertyName, string value)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(value);

        string? matchedName = MatchSupportedBuiltInProperty(propertyName);
        if (matchedName is null)
        {
            return UnsupportedBuiltInPropertyError(propertyName);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? property = null;
            try
            {
                property = ctx.Presentation.BuiltInDocumentProperties[matchedName];
                property.Value = value;

                return new PresentationOperationResult
                {
                    Success = true,
                    PresentationPath = batch.PresentationPath,
                    PropertyName = matchedName,
                    PropertyValue = (string)property.Value
                };
            }
            finally
            {
                if (property != null)
                {
                    ComUtilities.Release(ref property!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult GetDocumentProperty(IPresentationBatch batch, string propertyName)
    {
        ArgumentNullException.ThrowIfNull(batch);

        string? matchedName = MatchSupportedBuiltInProperty(propertyName);
        if (matchedName is null)
        {
            return UnsupportedBuiltInPropertyError(propertyName);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? property = null;
            try
            {
                property = ctx.Presentation.BuiltInDocumentProperties[matchedName];

                return new PresentationOperationResult
                {
                    Success = true,
                    PresentationPath = batch.PresentationPath,
                    PropertyName = matchedName,
                    PropertyValue = (string)property.Value
                };
            }
            finally
            {
                if (property != null)
                {
                    ComUtilities.Release(ref property!);
                }
            }
        });
    }

    private static string? MatchSupportedBuiltInProperty(string propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return null;
        }

        foreach (string candidate in SupportedBuiltInProperties)
        {
            if (string.Equals(candidate, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }
        }

        return null;
    }

    private static PresentationOperationResult UnsupportedBuiltInPropertyError(string propertyName) => new()
    {
        Success = false,
        ErrorMessage = $"'{propertyName}' is not a supported built-in document property. " +
                       $"Expected one of: {string.Join(", ", SupportedBuiltInProperties)}."
    };

    /// <inheritdoc/>
    public PresentationOperationResult SetCustomProperty(IPresentationBatch batch, string propertyName, string value)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(value);

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = "A custom property name is required."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var custom = ctx.Presentation.CustomDocumentProperties;

            // PowerPoint's CustomDocumentProperties collection has no TryGetValue/Contains
            // helper — an ArgumentException from the name-indexed lookup is the documented way
            // to detect "not present yet" (verified live via COM spike), so this upsert pattern
            // is a normal existence check, not suppression of an unexpected failure (Rule 1b).
            dynamic? existing = null;
            try
            {
                try
                {
                    existing = custom[propertyName];
                    existing.Value = value;
                }
                catch (ArgumentException)
                {
                    custom.Add(propertyName, false, MsoPropertyTypeString, value);
                }
            }
            finally
            {
                if (existing != null)
                {
                    ComUtilities.Release(ref existing!);
                }
            }

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath,
                PropertyName = propertyName,
                PropertyValue = value
            };
        });
    }

    /// <summary>
    /// <c>MsoDocProperties.msoPropertyTypeString</c> — used directly as an <c>int</c> to avoid
    /// pulling in the full <c>Microsoft.Office.Core</c> interop surface for a single enum value.
    /// </summary>
    private const int MsoPropertyTypeString = 4;

    /// <inheritdoc/>
    public PresentationOperationResult GetCustomProperty(IPresentationBatch batch, string propertyName)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = "A custom property name is required."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var custom = ctx.Presentation.CustomDocumentProperties;

            dynamic? existing = null;
            try
            {
                try
                {
                    existing = custom[propertyName];

                    return new PresentationOperationResult
                    {
                        Success = true,
                        PresentationPath = batch.PresentationPath,
                        PropertyName = propertyName,
                        PropertyValue = (string)existing.Value
                    };
                }
                catch (ArgumentException)
                {
                    return new PresentationOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"No custom property named '{propertyName}' was found."
                    };
                }
            }
            finally
            {
                if (existing != null)
                {
                    ComUtilities.Release(ref existing!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult RemoveCustomProperty(IPresentationBatch batch, string propertyName)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return new PresentationOperationResult
            {
                Success = false,
                ErrorMessage = "A custom property name is required."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var custom = ctx.Presentation.CustomDocumentProperties;

            dynamic? existing = null;
            try
            {
                try
                {
                    existing = custom[propertyName];
                    existing.Delete();

                    return new PresentationOperationResult
                    {
                        Success = true,
                        PresentationPath = batch.PresentationPath,
                        PropertyName = propertyName
                    };
                }
                catch (ArgumentException)
                {
                    return new PresentationOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"No custom property named '{propertyName}' was found."
                    };
                }
            }
            finally
            {
                if (existing != null)
                {
                    ComUtilities.Release(ref existing!);
                }
            }
        });
    }
}
