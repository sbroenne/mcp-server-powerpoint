using System.Text.Json;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that route to a Core service category.
/// </summary>
/// <remarks>
/// Structural port of Sbroenne.ExcelMcp.CLI.Infrastructure.ServiceCommandBase&lt;TSettings&gt;,
/// scoped down to what the generated CLI command classes need to compile. PowerPointMcp has NO
/// out-of-process Service/daemon (see squad decision: "SESSION OWNERSHIP — DIVERGE FROM EXCEL FOR
/// THE MVP" — an in-process <c>PresentationSessionRegistry</c> is used instead, owned by the
/// MCP server host). Excel's version dispatches over a named-pipe daemon (<c>DaemonAutoStart</c> +
/// <c>ServiceBridge</c>); porting that is out of scope for this pilot. This base class still
/// performs the same session/action validation and calls <see cref="Route"/> to prove the
/// generated routing code round-trips correctly, but reports a clear "not yet wired" error instead
/// of dispatching, since PowerPointMcp.CLI has no long-lived process to hold an open
/// IPresentationBatch across separate CLI invocations. Wiring a real dispatch target (e.g. an
/// in-process session store shared with a future daemon, or direct Core calls that open+close a
/// batch per invocation) is a follow-up, not part of proving the source-generator pipeline.
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

    /// <summary>Whether this command requires a session ID. Default is true.</summary>
    protected virtual bool RequiresSession => true;

    /// <inheritdoc/>
    protected sealed override Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        var sessionId = GetSessionId(settings);
        if (RequiresSession && string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(WriteError("Session ID is required. Use --session <id>"));
        }

        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            return Task.FromResult(WriteError("Action is required."));
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            return Task.FromResult(WriteError($"Invalid action '{action}'. Valid actions: {validList}"));
        }

        // Route and validate parameters (proves the generated RouteFromSettings/RouteCliArgs code
        // — including ParameterTransforms validation — works end-to-end).
        string command;
        object? args;
        try
        {
            (command, args) = Route(settings, action);
        }
        catch (ArgumentException ex)
        {
            return Task.FromResult(WriteError(ex.Message));
        }

        // No out-of-process Service exists yet (deliberate MVP divergence from Excel) — nothing to
        // dispatch `command`/`args` to. Report this clearly instead of pretending to execute.
        return Task.FromResult(WriteError(
            $"'{command}' routed successfully but PowerPointMcp.CLI has no execution backend yet " +
            "(no out-of-process Service — MCP server owns sessions in-process). " +
            "Use the MCP server tools for this operation."));
    }

    private static int WriteError(string message)
    {
        Console.Error.WriteLine(JsonSerializer.Serialize(new { success = false, error = message }));
        return 1;
    }
}
