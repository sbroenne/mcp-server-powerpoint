using System.Text.Json;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that route to a Core service category.
/// </summary>
/// <remarks>
/// Structural port of Sbroenne.ExcelMcp.CLI.Infrastructure.ServiceCommandBase&lt;TSettings&gt;,
/// adapted to PowerPointMcp's stateless CLI model. PowerPointMcp has NO out-of-process
/// Service/daemon (squad decision, 2026-07-06: "DROP the Service — functional overkill").
/// Excel's version dispatches an action string over a named-pipe daemon (<c>DaemonAutoStart</c> +
/// <c>ServiceBridge</c>) that holds an open document session across SEPARATE CLI process
/// invocations. PowerPointMcp.CLI has no such long-lived process, so it cannot carry a
/// "--session &lt;id&gt;" across invocations the way Excel's CLI does.
///
/// Instead, each invocation is fully self-contained (mirrors the hand-written placeholder
/// Program.cs, which already proves this model): open the target .pptx file directly via
/// <c>PresentationSession.BeginBatch</c>, execute exactly one Core action, save, and dispose —
/// no persistent session ID needed because there is no persistent process to hold one. The
/// generated per-category <c>Execute</c> override does this open/dispatch/save/dispose work
/// using the generated <c>ServiceRegistry.{Category}.DispatchToCore</c> method, driven by the
/// "--file &lt;path&gt;" (<c>DocumentPath</c>) option in generated CLI settings instead of Excel's
/// "--session &lt;id&gt;".
/// </remarks>
internal abstract class ServiceCommandBase<TSettings> : AsyncCommand<TSettings>
    where TSettings : CommandSettings
{
    /// <summary>Gets the action from settings.</summary>
    protected abstract string? GetAction(TSettings settings);

    /// <summary>Gets the valid actions for this command.</summary>
    protected abstract IReadOnlyList<string> ValidActions { get; }

    /// <summary>
    /// Routes the action to a (command, args) pair using the generated
    /// <c>ServiceRegistry.{Category}.RouteFromSettings</c>. Also performs parameter validation
    /// (e.g. required-parameter checks) — an <see cref="ArgumentException"/> here means bad
    /// CLI input, not a Core/COM failure.
    /// </summary>
    protected abstract (string command, object? args) Route(TSettings settings, string action);

    /// <summary>
    /// Executes the routed action against a real PowerPoint file for this single CLI invocation:
    /// opens the target file (if the action needs one), calls the generated
    /// <c>ServiceRegistry.{Category}.DispatchToCore</c>, saves, and disposes — all within this one
    /// call. Generated per category by <c>CliSettingsGenerator</c>. Returns the JSON result
    /// payload to print to stdout.
    /// </summary>
    protected abstract string Execute(TSettings settings, string action, object? args);

    /// <inheritdoc/>
    protected sealed override Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
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

        object? args;
        try
        {
            (_, args) = Route(settings, action);
        }
        catch (ArgumentException ex)
        {
            return Task.FromResult(WriteError(ex.Message));
        }

        // CLI process boundary (mirrors PowerPointToolsBase.ExecuteToolAction at the MCP layer):
        // this is the ONE place in the CLI allowed to catch-and-report — Core itself never
        // suppresses exceptions (Rule 1b).
        try
        {
            var resultJson = Execute(settings, action, args);
            Console.WriteLine(resultJson);
            return Task.FromResult(0);
        }
        catch (Exception ex)
        {
            return Task.FromResult(WriteError(ex.Message));
        }
    }

    private static int WriteError(string message)
    {
        Console.Error.WriteLine(JsonSerializer.Serialize(new { success = false, error = message }));
        return 1;
    }
}
