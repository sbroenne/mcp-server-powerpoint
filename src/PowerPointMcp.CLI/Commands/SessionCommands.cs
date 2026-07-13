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

/// <summary>Settings for the "session apply-template" command.</summary>
internal sealed class SessionApplyTemplateSettings : CommandSettings
{
    [CommandArgument(0, "<SESSION_ID>")]
    [Description("Session id returned by 'session open'/'session create'.")]
    public string SessionId { get; init; } = string.Empty;

    [CommandArgument(1, "<TEMPLATE_PATH>")]
    [Description("Full path to a .potx/.potm/.pot template file (or a .pptx/.pptm presentation used as a template source).")]
    public string TemplatePath { get; init; } = string.Empty;
}

/// <summary>Applies a template's masters/theme/layouts to the open presentation, preserving slide content.</summary>
internal sealed class SessionApplyTemplateCommand : AsyncCommand<SessionApplyTemplateSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionApplyTemplateSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.apply-template",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { templatePath = settings.TemplatePath }, ServiceProtocol.JsonOptions),
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

/// <summary>Reads the design/theme name currently applied to the open presentation.</summary>
internal sealed class SessionGetThemeNameCommand : AsyncCommand<SessionSaveSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionSaveSettings settings, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.get-theme-name",
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

/// <summary>Settings shared by the document/custom property "get" commands.</summary>
internal sealed class SessionPropertyGetSettings : CommandSettings
{
    [CommandArgument(0, "<SESSION_ID>")]
    [Description("Session id returned by 'session open'/'session create'.")]
    public string SessionId { get; init; } = string.Empty;

    [CommandArgument(1, "<PROPERTY_NAME>")]
    [Description("Document property name (built-in: Title, Subject, Author, Keywords, Comments, Category, Manager, Company; or any custom name).")]
    public string PropertyName { get; init; } = string.Empty;
}

/// <summary>Settings shared by the document/custom property "set" commands.</summary>
internal sealed class SessionPropertySetSettings : CommandSettings
{
    [CommandArgument(0, "<SESSION_ID>")]
    [Description("Session id returned by 'session open'/'session create'.")]
    public string SessionId { get; init; } = string.Empty;

    [CommandArgument(1, "<PROPERTY_NAME>")]
    [Description("Document property name (built-in: Title, Subject, Author, Keywords, Comments, Category, Manager, Company; or any custom name).")]
    public string PropertyName { get; init; } = string.Empty;

    [CommandArgument(2, "<VALUE>")]
    [Description("The new property value.")]
    public string Value { get; init; } = string.Empty;
}

/// <summary>Sets a built-in document metadata property on the open presentation.</summary>
internal sealed class SessionSetDocumentPropertyCommand : AsyncCommand<SessionPropertySetSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionPropertySetSettings settings, CancellationToken cancellationToken)
        => await SessionCommandHelpers.SendPropertyCommandAsync("session.set-document-property", settings.SessionId, settings.PropertyName, settings.Value, cancellationToken);
}

/// <summary>Reads a built-in document metadata property from the open presentation.</summary>
internal sealed class SessionGetDocumentPropertyCommand : AsyncCommand<SessionPropertyGetSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionPropertyGetSettings settings, CancellationToken cancellationToken)
        => await SessionCommandHelpers.SendPropertyCommandAsync("session.get-document-property", settings.SessionId, settings.PropertyName, value: null, cancellationToken);
}

/// <summary>Creates or updates a custom (user-defined) document property on the open presentation.</summary>
internal sealed class SessionSetCustomPropertyCommand : AsyncCommand<SessionPropertySetSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionPropertySetSettings settings, CancellationToken cancellationToken)
        => await SessionCommandHelpers.SendPropertyCommandAsync("session.set-custom-property", settings.SessionId, settings.PropertyName, settings.Value, cancellationToken);
}

/// <summary>Reads a custom (user-defined) document property from the open presentation.</summary>
internal sealed class SessionGetCustomPropertyCommand : AsyncCommand<SessionPropertyGetSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionPropertyGetSettings settings, CancellationToken cancellationToken)
        => await SessionCommandHelpers.SendPropertyCommandAsync("session.get-custom-property", settings.SessionId, settings.PropertyName, value: null, cancellationToken);
}

/// <summary>Removes a custom (user-defined) document property from the open presentation.</summary>
internal sealed class SessionRemoveCustomPropertyCommand : AsyncCommand<SessionPropertyGetSettings>
{
    /// <inheritdoc/>
    protected override async Task<int> ExecuteAsync(CommandContext context, SessionPropertyGetSettings settings, CancellationToken cancellationToken)
        => await SessionCommandHelpers.SendPropertyCommandAsync("session.remove-custom-property", settings.SessionId, settings.PropertyName, value: null, cancellationToken);
}

/// <summary>Shared daemon round-trip logic for the document/custom-property session commands.</summary>
internal static class SessionCommandHelpers
{
    public static async Task<int> SendPropertyCommandAsync(string command, string sessionId, string propertyName, string? value, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { propertyName, value }, ServiceProtocol.JsonOptions),
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
