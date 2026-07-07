using System.ComponentModel;
using System.Text.Json;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Commands;

/// <summary>Settings shared by "service" subcommands that need the pipe name.</summary>
internal class ServicePipeSettings : CommandSettings
{
    [CommandOption("--pipe-name <PIPE_NAME>")]
    [Description("Override the daemon's named pipe (defaults to a per-user pipe name).")]
    public string? PipeName { get; init; }

    /// <summary>Gets the effective pipe name, defaulting to the shared per-user CLI pipe.</summary>
    public string ResolvedPipeName => string.IsNullOrWhiteSpace(PipeName) ? DaemonAutoStart.GetPipeName() : PipeName;
}

/// <summary>Starts the CLI daemon if it isn't already running (auto-starts, then reports status).</summary>
internal sealed class ServiceStartCommand : AsyncCommand<ServicePipeSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, ServicePipeSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest { Command = "service.status", Source = "cli" }, cancellationToken);
        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>Settings for "service stop".</summary>
internal sealed class ServiceStopSettings : ServicePipeSettings
{
    [CommandOption("--force")]
    [Description("Force-kill the daemon process if a graceful RPC shutdown doesn't respond.")]
    public bool Force { get; init; }
}

/// <summary>Stops the running CLI daemon (graceful RPC shutdown, or force-kill with --force).</summary>
internal sealed class ServiceStopCommand : AsyncCommand<ServiceStopSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, ServiceStopSettings settings, CancellationToken cancellationToken)
    {
        var pipeName = settings.ResolvedPipeName;
        using var client = new ServiceClient(pipeName);

        if (await client.PingAsync(cancellationToken))
        {
            var response = await client.SendAsync(new ServiceRequest { Command = "service.shutdown", Source = "cli" }, cancellationToken);
            if (response.Success)
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Daemon shutdown requested." }, ServiceProtocol.JsonOptions));
                return 0;
            }

            if (!settings.Force)
            {
                return CliErrorOutput.WriteServiceError(response);
            }
        }
        else if (!settings.Force)
        {
            return CliErrorOutput.WriteError("Daemon is not running.");
        }

        if (settings.Force && DaemonProcessTracker.TryForceStopTrackedDaemon(pipeName))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Daemon process force-killed." }, ServiceProtocol.JsonOptions));
            return 0;
        }

        return CliErrorOutput.WriteError("Daemon is not running or could not be stopped.");
    }
}

/// <summary>Reports whether the CLI daemon is running, plus its session count and uptime.</summary>
internal sealed class ServiceStatusCommand : AsyncCommand<ServicePipeSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, ServicePipeSettings settings, CancellationToken cancellationToken)
    {
        var pipeName = settings.ResolvedPipeName;
        using var client = new ServiceClient(pipeName);

        if (!await client.PingAsync(cancellationToken))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, running = false }, ServiceProtocol.JsonOptions));
            return 0;
        }

        var response = await client.SendAsync(new ServiceRequest { Command = "service.status", Source = "cli" }, cancellationToken);
        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}
