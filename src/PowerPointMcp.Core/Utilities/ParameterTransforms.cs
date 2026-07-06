using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Utilities;

/// <summary>
/// Shared parameter transformation utilities used by MCP, CLI, and generated code.
/// These provide consistent handling of common patterns across all entry points.
/// </summary>
/// <remarks>
/// Ported from Sbroenne.ExcelMcp.Core.Utilities.ParameterTransforms — only the
/// domain-agnostic members are included here (no Excel-specific option builders
/// like FindOptions/ReplaceOptions/PowerQueryLoadMode, since those types don't
/// exist in PowerPoint Core). Add more members as later domains need them.
/// </remarks>
public static class ParameterTransforms
{
    /// <summary>
    /// Validates that a required parameter is not null, empty, or whitespace.
    /// </summary>
    /// <param name="value">The parameter value to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="actionName">Name of the action for error messages</param>
    /// <exception cref="ArgumentException">Thrown when value is null, empty, or whitespace</exception>
    public static void RequireNotEmpty(string? value, string parameterName, string actionName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"{parameterName} is required for {actionName} action", parameterName);
        }
    }

    /// <summary>
    /// Validates that a required parameter is not null, empty, or whitespace, returning the value if valid.
    /// </summary>
    /// <param name="value">The parameter value to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="actionName">Name of the action for error messages</param>
    /// <returns>The validated non-null value</returns>
    /// <exception cref="ArgumentException">Thrown when value is null, empty, or whitespace</exception>
    public static string RequireNotEmptyReturn(string? value, string parameterName, string actionName)
    {
        RequireNotEmpty(value, parameterName, actionName);
        return value!;
    }

    /// <summary>
    /// Resolves a value that can come from either a direct string or a file path.
    /// If filePath is provided and exists, reads file content. Otherwise returns directValue.
    /// </summary>
    /// <param name="directValue">The direct string value (e.g., text inline)</param>
    /// <param name="filePath">Optional path to a file containing the value</param>
    /// <returns>The resolved value (file content or direct value)</returns>
    public static string? ResolveFileOrValue(string? directValue, string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}", filePath);
            }
            return File.ReadAllText(filePath);
        }
        return directValue;
    }

    /// <summary>
    /// Parses a timeout value supplied via CLI.
    /// Plain numeric values are interpreted as seconds; TimeSpan-formatted values are preserved.
    /// Returns null for null/empty input.
    /// </summary>
    /// <param name="value">Timeout text from CLI</param>
    /// <param name="parameterName">Parameter name for error messages</param>
    /// <returns>Parsed timeout or null</returns>
    /// <exception cref="FormatException">Thrown when the timeout cannot be parsed</exception>
    public static TimeSpan? ParseTimeSpanOrSeconds(string? value, string parameterName = "timeout")
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var seconds))
            return TimeSpan.FromSeconds(seconds);

        if (TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out var parsed)
            || TimeSpan.TryParse(value, CultureInfo.CurrentCulture, out parsed))
        {
            return parsed;
        }

        throw new FormatException(
            $"Invalid {parameterName} value '{value}'. Use seconds (for example 600) or TimeSpan format (for example 00:10:00).");
    }
}
