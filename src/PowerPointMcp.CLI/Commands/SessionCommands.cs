using System.ComponentModel;
using System.Text.Json;
using Sbroenne.PowerPointMcp.CLI.Infrastructure;
using Sbroenne.PowerPointMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Commands;

/// <summary>Settings shared by the "session open"/"session create" commands.</summary>
internal sealed class SessionOpenSettings : CommandSettings
{
    [CommandArgument(0, "<FILE_PATH>")]
    [Description("Full path to the .pptx/.pptm presentation file.")]
    public string FilePath { get; init; } = string.Empty;
}

/// <summary>Opens an existing presentation and returns a session id for subsequent commands.</summary>
internal sealed class SessionOpenCommand : AsyncCommand<SessionOpenSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionOpenSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath }, ServiceProtocol.JsonOptions),
            Source = "cli"
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>Creates a new presentation and returns a session id for subsequent commands.</summary>
internal sealed class SessionCreateCommand : AsyncCommand<SessionOpenSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionOpenSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath }, ServiceProtocol.JsonOptions),
            Source = "cli"
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>Settings for the "session close" command.</summary>
internal sealed class SessionCloseSettings : CommandSettings
{
    [CommandArgument(0, "<SESSION_ID>")]
    [Description("Session id returned by 'session open'/'session create'.")]
    public string SessionId { get; init; } = string.Empty;

    [CommandOption("--save")]
    [Description("Save the presentation before closing it.")]
    public bool Save { get; init; }
}

/// <summary>Closes a session, disposing its PowerPoint batch (in the background).</summary>
internal sealed class SessionCloseCommand : AsyncCommand<SessionCloseSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionCloseSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { save = settings.Save }, ServiceProtocol.JsonOptions),
            Source = "cli"
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>Settings for the "session save" command.</summary>
internal sealed class SessionSaveSettings : CommandSettings
{
    [CommandArgument(0, "<SESSION_ID>")]
    [Description("Session id returned by 'session open'/'session create'.")]
    public string SessionId { get; init; } = string.Empty;
}

/// <summary>Saves the presentation open in a session without closing it.</summary>
internal sealed class SessionSaveCommand : AsyncCommand<SessionSaveSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionSaveSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = settings.SessionId,
            Source = "cli"
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>Lists every session currently open in the daemon.</summary>
internal sealed class SessionListCommand : AsyncCommand
{
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(2);

    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        using var client = new ServiceClient(pipeName, connectTimeout: CommandTimeout, requestTimeout: CommandTimeout);

        try
        {
            var response = await client.SendAsync(new ServiceRequest { Command = "session.list", Source = "cli" }, cancellationToken);
            if (response.Success)
            {
                Console.WriteLine(response.Result ?? JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, ServiceProtocol.JsonOptions));
                return 0;
            }

            return CliErrorOutput.WriteServiceError(response);
        }
        catch (Exception)
        {
            // Daemon not running — there are, by definition, no open sessions.
            Console.WriteLine(JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, ServiceProtocol.JsonOptions));
            return 0;
        }
    }
}
