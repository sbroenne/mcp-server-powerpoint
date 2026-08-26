using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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

        return batch.Execute((ctx, ct) =>
        {
            // SetFinal saves pending edits before enabling Final. PowerPoint then persists the
            // flag and makes the presentation read-only, so a second Save would fail with
            // "Presentation cannot be modified". Save-on-close is therefore a successful no-op.
            if (!ctx.Presentation.Final)
            {
                ctx.Presentation.Save();
            }

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath
            };
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult SaveAs(
        IPresentationBatch batch,
        string targetPath,
        PresentationSaveFormat format = PresentationSaveFormat.Auto,
        bool overwrite = false)
    {
        ArgumentNullException.ThrowIfNull(batch);

        var pathValidation = ValidateOutputPath(batch, targetPath, overwrite);
        if (pathValidation.Error != null)
        {
            return pathValidation.Error;
        }

        string normalizedPath = pathValidation.Path!;
        PresentationSaveFormat? resolvedFormat = ResolveSaveFormat(normalizedPath, format);
        if (resolvedFormat == null)
        {
            return ValidationError(
                batch,
                $"Cannot infer a supported presentation format from extension '{Path.GetExtension(normalizedPath)}'. Supported extensions: .pptx, .pptm, .ppt.");
        }

        var extensionError = ValidateSaveExtension(batch, normalizedPath, resolvedFormat);
        if (extensionError != null)
        {
            return extensionError;
        }

        var result = batch.Execute((ctx, ct) =>
        {
            // PIA gap: the restored typed SaveAs method exposes an optional Office.MsoTriState
            // parameter, but this project intentionally does not reference office.dll. Keep late
            // binding limited to this invocation while still passing typed PpSaveAsFileType.
            ((dynamic)ctx.Presentation).SaveAs(
                normalizedPath,
                ToPowerPointFileType(resolvedFormat.Value));

            batch.UpdatePresentationPath(normalizedPath);
            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = normalizedPath
            };
        });

        return result;
    }

    /// <inheritdoc/>
    public PresentationOperationResult SaveCopyAs(
        IPresentationBatch batch,
        string targetPath,
        bool overwrite = false)
    {
        ArgumentNullException.ThrowIfNull(batch);

        var pathValidation = ValidateOutputPath(batch, targetPath, overwrite);
        if (pathValidation.Error != null)
        {
            return pathValidation.Error;
        }

        string normalizedPath = pathValidation.Path!;
        string outputExtension = Path.GetExtension(normalizedPath);
        PresentationSaveFormat? format = ResolveSaveFormat(
            normalizedPath,
            PresentationSaveFormat.Auto);
        if (format == null)
        {
            return ValidationError(
                batch,
                $"Save Copy As does not support presentation extension '{outputExtension}'. Supported extensions: .pptx, .pptm, .ppt.");
        }

        string writePath = GetWritePath(normalizedPath, overwrite);
        try
        {
            var result = batch.Execute((ctx, ct) =>
            {
                string currentExtension = Path.GetExtension(batch.PresentationPath);
                if (!string.Equals(currentExtension, outputExtension, StringComparison.OrdinalIgnoreCase))
                {
                    return ValidationError(
                        batch,
                        $"Save Copy As preserves the current presentation format. Output extension must be '{currentExtension}'.");
                }

                // PIA gap: the restored typed SaveCopyAs method exposes an optional
                // Office.MsoTriState parameter, but this project intentionally does not reference
                // office.dll. Keep late binding limited to this invocation while still passing
                // typed PpSaveAsFileType.
                ((dynamic)ctx.Presentation).SaveCopyAs(
                    writePath,
                    ToPowerPointFileType(format.Value));

                return new PresentationOperationResult
                {
                    Success = true,
                    PresentationPath = normalizedPath
                };
            });

            if (!result.Success)
            {
                return result;
            }

            CommitOutput(writePath, normalizedPath);
            return result;
        }
        finally
        {
            DeleteTemporaryOutput(writePath, normalizedPath);
        }
    }

    private static (string? Path, PresentationOperationResult? Error) ValidateOutputPath(
        IPresentationBatch batch,
        string targetPath,
        bool overwrite)
    {
        if (string.IsNullOrWhiteSpace(targetPath))
        {
            return (null, ValidationError(batch, "A destination path is required."));
        }

        string normalizedPath;
        try
        {
            normalizedPath = Path.GetFullPath(targetPath);
        }
        catch (ArgumentException)
        {
            return (null, ValidationError(batch, $"Invalid destination path: '{targetPath}'."));
        }
        catch (NotSupportedException)
        {
            return (null, ValidationError(batch, $"Invalid destination path: '{targetPath}'."));
        }
        catch (PathTooLongException)
        {
            return (null, ValidationError(batch, $"Destination path is too long: '{targetPath}'."));
        }

        string? directory = Path.GetDirectoryName(normalizedPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            return (null, ValidationError(batch, $"Destination directory does not exist: '{directory}'."));
        }

        if (Directory.Exists(normalizedPath))
        {
            return (null, ValidationError(batch, $"Destination path is a directory: '{normalizedPath}'."));
        }

        if (File.Exists(normalizedPath) && !overwrite)
        {
            return (null, ValidationError(batch, $"Destination file already exists: '{normalizedPath}'. Set overwrite=true to replace it."));
        }

        return (normalizedPath, null);
    }

    private static PresentationSaveFormat? ResolveSaveFormat(
        string outputPath,
        PresentationSaveFormat format)
    {
        if (format != PresentationSaveFormat.Auto)
        {
            return Enum.IsDefined(format) ? format : null;
        }

        return Path.GetExtension(outputPath).ToLowerInvariant() switch
        {
            ".pptx" => PresentationSaveFormat.Pptx,
            ".pptm" => PresentationSaveFormat.Pptm,
            ".ppt" => PresentationSaveFormat.Ppt,
            _ => null
        };
    }

    private static PresentationOperationResult? ValidateSaveExtension(
        IPresentationBatch batch,
        string outputPath,
        PresentationSaveFormat? format)
    {
        string expectedExtension = format switch
        {
            PresentationSaveFormat.Pptx => ".pptx",
            PresentationSaveFormat.Pptm => ".pptm",
            PresentationSaveFormat.Ppt => ".ppt",
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Save As format.")
        };

        return string.Equals(
            Path.GetExtension(outputPath),
            expectedExtension,
            StringComparison.OrdinalIgnoreCase)
            ? null
            : ValidationError(
                batch,
                $"Save As format '{format}' requires the '{expectedExtension}' file extension.");
    }

    private static PowerPoint.PpSaveAsFileType ToPowerPointFileType(PresentationSaveFormat format)
    {
        return format switch
        {
            PresentationSaveFormat.Pptx => PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
            PresentationSaveFormat.Pptm => PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled,
            PresentationSaveFormat.Ppt => PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Save As format.")
        };
    }

    private static string GetWritePath(string normalizedPath, bool overwrite)
    {
        if (!overwrite || !File.Exists(normalizedPath))
        {
            return normalizedPath;
        }

        string directory = Path.GetDirectoryName(normalizedPath)!;
        string fileName = Path.GetFileNameWithoutExtension(normalizedPath);
        string extension = Path.GetExtension(normalizedPath);
        return Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}.tmp{extension}");
    }

    private static void CommitOutput(string writePath, string normalizedPath)
    {
        if (!string.Equals(writePath, normalizedPath, StringComparison.OrdinalIgnoreCase))
        {
            File.Move(writePath, normalizedPath, overwrite: true);
        }
    }

    private static void DeleteTemporaryOutput(string writePath, string normalizedPath)
    {
        if (!string.Equals(writePath, normalizedPath, StringComparison.OrdinalIgnoreCase) &&
            File.Exists(writePath))
        {
            File.Delete(writePath);
        }
    }

    private static PresentationOperationResult ValidationError(
        IPresentationBatch batch,
        string message)
        => new()
        {
            Success = false,
            ErrorMessage = message,
            PresentationPath = batch.PresentationPath
        };

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

            string? themeName = ReadFirstDesignName(ctx.Presentation);

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
            string? themeName = ReadFirstDesignName(ctx.Presentation);

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath,
                ThemeName = themeName
            };
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult GetFinal(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) => new PresentationOperationResult
        {
            Success = true,
            PresentationPath = batch.PresentationPath,
            IsFinal = ctx.Presentation.Final
        });
    }

    /// <inheritdoc/>
    public PresentationOperationResult SetFinal(IPresentationBatch batch, bool isFinal)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            if (isFinal && !ctx.Presentation.Final)
            {
                ctx.Presentation.Save();
            }

            ctx.Presentation.Final = isFinal;

            return new PresentationOperationResult
            {
                Success = true,
                PresentationPath = batch.PresentationPath,
                IsFinal = ctx.Presentation.Final
            };
        });
    }

    /// <summary>
    /// Reads the Name of the first Design in the presentation's Designs collection, if any,
    /// releasing both the intermediate DesignCollection and Design COM objects afterward.
    /// </summary>
    private static string? ReadFirstDesignName(PowerPoint.Presentation presentation)
    {
        dynamic? designs = null;
        dynamic? design = null;
        try
        {
            designs = presentation.Designs;
            if (designs.Count == 0) return null;

            design = designs[1];
            return (string)design.Name;
        }
        finally
        {
            if (design != null) ComUtilities.Release(ref design!);
            if (designs != null) ComUtilities.Release(ref designs!);
        }
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
            dynamic custom = ctx.Presentation.CustomDocumentProperties;

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
                ComUtilities.Release(ref custom!);
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
            dynamic custom = ctx.Presentation.CustomDocumentProperties;

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
                ComUtilities.Release(ref custom!);
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
            dynamic custom = ctx.Presentation.CustomDocumentProperties;

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
                ComUtilities.Release(ref custom!);
            }
        });
    }
}
