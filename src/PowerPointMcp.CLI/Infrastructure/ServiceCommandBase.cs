using System.Text.Json;
using Sbroenne.PowerPointMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that send requests to the PowerPointMCP CLI daemon.
/// Handles common session/action validation and dispatch over the named-pipe RPC transport.
/// </summary>
/// <remarks>
/// Structural port of Sbroenne.ExcelMcp.CLI.Infrastructure.ServiceCommandBase&lt;TSettings&gt;.
/// PowerPointMCP now HAS an out-of-process daemon (<c>PowerPointMcp.Service</c>) — see squad
/// decision 2026-07-07, which reverses the earlier "drop the Service — functional overkill" call
/// (2026-07-06). The daemon holds one long-lived <c>IPresentationBatch</c> per session id across
/// separate CLI process invocations, so a caller only pays PowerPoint's ~90-150s launch/teardown
/// cost once per "session open"/"session create", not on every command. Each generated command
/// class supplies <see cref="GetSessionId"/>/<see cref="GetAction"/>/<see cref="ValidActions"/>/
/// <see cref="Route"/>; this base class validates them, then hands the routed
/// <c>(command, args)</c> pair to <see cref="DaemonAutoStart.EnsureAndConnectAsync"/> +
/// <see cref="ServiceClient.SendAsync"/>, auto-starting the daemon on first use.
/// </remarks>
internal abstract class ServiceCommandBase<TSettings> : AsyncCommand<TSettings>
    where TSettings : CommandSettings
{
    /// <summary>Gets the session ID from settings.</summary>
    protected abstract string? GetSessionId(TSettings settings);

    /// <summary>Gets the action from settings.</summary>
    protected abstract string? GetAction(TSettings settings);

    /// <summary>Gets the valid actions for this command.</summary>
    protected abstract IReadOnlyList<string> ValidActions { get; }

    /// <summary>Routes the action to a (command, args) pair using the generated ServiceRegistry.</summary>
    protected abstract (string command, object? args) Route(TSettings settings, string action);

    /// <summary>
    /// Whether this command requires a session ID. Default is true. Override to return false for
    /// commands that don't need a session.
    /// </summary>
    protected virtual bool RequiresSession => true;

    /// <inheritdoc/>
    protected sealed override async Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        var sessionId = GetSessionId(settings);
        if (RequiresSession && string.IsNullOrWhiteSpace(sessionId))
        {
            return CliErrorOutput.WriteError("Session ID is required. Use --session <id> (see 'pptcli session open'/'pptcli session create').");
        }

        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            return CliErrorOutput.WriteError("Action is required.");
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            return CliErrorOutput.WriteError($"Invalid action '{action}'. Valid actions: {validList}");
        }

        string command;
        object? args;
        try
        {
            (command, args) = Route(settings, action);
        }
        catch (ArgumentException ex)
        {
            return CliErrorOutput.WriteError(ex.Message);
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null,
            Source = "cli"
        }, cancellationToken);

        var outputPath = settings.GetType().GetProperty("OutputPath")?.GetValue(settings) as string;

        if (response.Success)
        {
            var result = !string.IsNullOrEmpty(response.Result)
                ? response.Result
                : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions);

            if (!string.IsNullOrEmpty(outputPath))
            {
                return WriteOutputToFile(result, outputPath);
            }

            Console.WriteLine(result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }

    /// <summary>
    /// Writes the result to a file. For image results containing base64 data, decodes and writes
    /// the binary image. Otherwise writes the JSON text.
    /// </summary>
    private static int WriteOutputToFile(string result, string outputPath)
    {
        try
        {
            var base64Data = TryExtractBase64Image(result);
            if (base64Data != null)
            {
                var imageBytes = Convert.FromBase64String(base64Data);
                File.WriteAllBytes(outputPath, imageBytes);

                var doc = JsonDocument.Parse(result);
                var metadata = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["outputPath"] = outputPath,
                    ["sizeBytes"] = imageBytes.Length
                };
                if (doc.RootElement.TryGetProperty("width", out var w)) metadata["width"] = w.GetInt32();
                if (doc.RootElement.TryGetProperty("height", out var h)) metadata["height"] = h.GetInt32();
                if (doc.RootElement.TryGetProperty("mimeType", out var m)) metadata["mimeType"] = m.GetString();
                Console.WriteLine(JsonSerializer.Serialize(metadata, ServiceProtocol.JsonOptions));
            }
            else
            {
                File.WriteAllText(outputPath, result);
                Console.WriteLine(JsonSerializer.Serialize(
                    new { success = true, outputPath },
                    ServiceProtocol.JsonOptions));
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(
                new { success = false, error = $"Failed to write output: {ex.Message}" },
                ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    /// <summary>Attempts to extract base64 image data from a JSON result.</summary>
    private static string? TryExtractBase64Image(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("imageBase64", out var imageElement))
            {
                return imageElement.GetString();
            }
        }
        catch (JsonException)
        {
            // Not valid JSON, can't extract image.
        }
        return null;
    }
}
