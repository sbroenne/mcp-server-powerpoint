using System.Collections.Concurrent;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Animation;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Export;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Layout;
using Sbroenne.PowerPointMcp.Core.Master;
using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.Core.SmartArt;
using Sbroenne.PowerPointMcp.Core.Table;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using Sbroenne.PowerPointMcp.Generated;
using Sbroenne.PowerPointMcp.Service.Rpc;
using StreamJsonRpc;

namespace Sbroenne.PowerPointMcp.Service;

/// <summary>
/// The PowerPointMCP CLI daemon. Holds a <see cref="PresentationSessionRegistry"/> (one
/// long-lived <see cref="IPresentationBatch"/> per session id, reused across separate CLI
/// process invocations) and dispatches incoming commands to Core via the generated
/// <c>ServiceRegistry.{Category}.DispatchToCore</c> methods.
/// </summary>
/// <remarks>
/// Ported from mcp-server-excel's <c>ExcelMcpService</c> (squad decision 2026-07-07, reversing
/// the 2026-07-06 "drop the Service" call): the daemon's whole reason for existing is to avoid
/// PowerPoint's ~90-150s launch/teardown cost on every CLI invocation — a session created by
/// "session open"/"session create" stays alive (PowerPoint process included) until "session
/// close"/idle-timeout/daemon shutdown, and every subsequent CLI command that references the
/// same session id reuses the same live PowerPoint process.
/// </remarks>
public sealed class PowerPointMcpService : IDisposable
{
    private readonly PresentationSessionRegistry _sessions = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly ConcurrentDictionary<Task, byte> _activeConnectionTasks = new();
    private readonly DateTime _startTime = DateTime.UtcNow;
    private string _pipeName = "";
    private TimeSpan? _idleTimeout;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private bool _disposed;

    private readonly PresentationCommands _presentationCommands = new();
    private readonly SlideCommands _slideCommands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextFrameCommands _textFrameCommands = new();
    private readonly TableCommands _tableCommands = new();
    private readonly ChartCommands _chartCommands = new();
    private readonly ExportCommands _exportCommands = new();
    private readonly ImageCommands _imageCommands = new();
    private readonly NotesCommands _notesCommands = new();
    private readonly LayoutCommands _layoutCommands = new();
    private readonly MasterCommands _masterCommands = new();
    private readonly AnimationCommands _animationCommands = new();
    private readonly SmartArtCommands _smartArtCommands = new();

    /// <summary>Gets the UTC time this daemon instance started.</summary>
    public DateTime StartTime => _startTime;

    /// <summary>Gets the number of currently open presentation sessions.</summary>
    public int SessionCount => _sessions.List().Count;

    /// <summary>
    /// Gets the session registry this daemon instance owns. Exposed so the MCP server can host a
    /// <see cref="PowerPointMcpService"/> in-process and hand this SAME registry instance to its
    /// hand-written session-lifecycle tools (<c>PresentationTools</c>) via dependency injection —
    /// mirroring mcp-server-excel's architecture, where one shared Service class is consumed two
    /// ways: in-process (direct calls, no pipe) by the MCP server, and via
    /// named-pipe/StreamJsonRpc by the separate CLI daemon process. The generated domain MCP
    /// tools (Slide, Shape, TextFrame, Table, Chart, Image, Notes, Layout, Master, Export) DO call
    /// <see cref="ProcessAsync"/> in-process via <c>ServiceBridge.ForwardToService</c> — only the
    /// hand-written <c>PresentationTools</c> bypass it and use <see cref="Sessions"/> directly.
    /// </summary>
    public PresentationSessionRegistry Sessions => _sessions;

    /// <summary>
    /// Runs the daemon, listening for commands on the named pipe. Blocks until shutdown is
    /// requested via <see cref="RequestShutdown"/> (explicit "service.shutdown" RPC or idle
    /// timeout).
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    /// <param name="idleTimeout">Shuts the daemon down after this much idle time with no open
    /// sessions. Null disables the idle monitor.</param>
    public async Task RunAsync(string pipeName, TimeSpan? idleTimeout = null)
    {
        _pipeName = pipeName;
        _idleTimeout = idleTimeout;
        await RunPipeServerAsync(_shutdownCts.Token);
    }

    /// <summary>Signals the pipe accept loop to stop, ending <see cref="RunAsync"/>.</summary>
    public void RequestShutdown() => _shutdownCts.Cancel();

    private void RequestShutdownAfterResponse()
    {
        _ = Task.Run(async () =>
        {
            await Task.Delay(100, CancellationToken.None);
            RequestShutdown();
        }, CancellationToken.None);
    }

    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);

    /// <summary>Records client activity to keep the idle timeout monitor alive.</summary>
    internal void RecordActivity() => _lastActivityTime = DateTime.UtcNow;

    private async Task RunPipeServerAsync(CancellationToken cancellationToken)
    {
        using var connectionLimit = new SemaphoreSlim(10, 10);

        if (_idleTimeout.HasValue)
        {
            _ = Task.Run(() => MonitorIdleTimeoutAsync(cancellationToken), cancellationToken);
        }

        var currentBackoff = InitialBackoff;

        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = ServiceSecurity.CreateSecureServer(_pipeName);
                await server.WaitForConnectionAsync(cancellationToken);

                currentBackoff = InitialBackoff;
                _lastActivityTime = DateTime.UtcNow;

                var clientServer = server;
                server = null;

                var connectionTask = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync(CancellationToken.None);
                    try
                    {
                        var rpcTarget = new DaemonRpcTarget(this);
                        using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
                        await rpc.Completion;
                    }
                    catch (Exception)
                    {
                        // Connection-level failures are expected on client disconnect; nothing to
                        // recover here — the accept loop simply serves the next connection.
                    }
                    finally
                    {
                        connectionLimit.Release();
                        try { if (clientServer.IsConnected) clientServer.Disconnect(); } catch (Exception) { }
                        try { await clientServer.DisposeAsync(); } catch (Exception) { }
                    }
                }, CancellationToken.None);
                _activeConnectionTasks.TryAdd(connectionTask, 0);
                _ = connectionTask.ContinueWith(
                    completed => _activeConnectionTasks.TryRemove(completed, out _),
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception)
            {
                try { await Task.Delay(currentBackoff, cancellationToken); } catch (OperationCanceledException) { break; }
                currentBackoff = TimeSpan.FromMilliseconds(Math.Min(currentBackoff.TotalMilliseconds * 2, MaxBackoff.TotalMilliseconds));
            }
            finally
            {
                if (server != null)
                {
                    try { if (server.IsConnected) server.Disconnect(); } catch (Exception) { }
                    await server.DisposeAsync();
                }
            }
        }

        if (!_activeConnectionTasks.IsEmpty)
        {
            await Task.WhenAll(_activeConnectionTasks.Keys.Select(ObserveConnectionTaskAsync));
        }
    }

    private static async Task ObserveConnectionTaskAsync(Task connectionTask)
    {
        try { await connectionTask; } catch (Exception) { }
    }

    private async Task MonitorIdleTimeoutAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);

            if (_sessions.List().Count > 0)
            {
                _lastActivityTime = DateTime.UtcNow;
                continue;
            }

            var idleTime = DateTime.UtcNow - _lastActivityTime;
            if (idleTime >= _idleTimeout!.Value)
            {
                RequestShutdown();
                break;
            }
        }
    }

    /// <summary>Processes a single request. Used by the pipe RPC target.</summary>
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        try
        {
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            ServiceResponse response = category switch
            {
                "service" => HandleServiceCommand(action),
                "session" => HandleSessionCommand(action, request),
                "slide" => DispatchSimple<SlideAction>(action, request,
                    ServiceRegistry.Slide.TryParseAction,
                    (a, batch) => ServiceRegistry.Slide.DispatchToCore(_slideCommands, a, batch, request.Args)),
                "shape" => DispatchSimple<ShapeAction>(action, request,
                    ServiceRegistry.Shape.TryParseAction,
                    (a, batch) => ServiceRegistry.Shape.DispatchToCore(_shapeCommands, a, batch, request.Args)),
                "textframe" => DispatchSimple<TextFrameAction>(action, request,
                    ServiceRegistry.TextFrame.TryParseAction,
                    (a, batch) => ServiceRegistry.TextFrame.DispatchToCore(_textFrameCommands, a, batch, request.Args)),
                "table" => DispatchSimple<TableAction>(action, request,
                    ServiceRegistry.Table.TryParseAction,
                    (a, batch) => ServiceRegistry.Table.DispatchToCore(_tableCommands, a, batch, request.Args)),
                "chart" => DispatchSimple<ChartAction>(action, request,
                    ServiceRegistry.Chart.TryParseAction,
                    (a, batch) => ServiceRegistry.Chart.DispatchToCore(_chartCommands, a, batch, request.Args)),
                "export" => DispatchSimple<ExportAction>(action, request,
                    ServiceRegistry.Export.TryParseAction,
                    (a, batch) => ServiceRegistry.Export.DispatchToCore(_exportCommands, a, batch, request.Args)),
                "image" => DispatchSimple<ImageAction>(action, request,
                    ServiceRegistry.Image.TryParseAction,
                    (a, batch) => ServiceRegistry.Image.DispatchToCore(_imageCommands, a, batch, request.Args)),
                "notes" => DispatchSimple<NotesAction>(action, request,
                    ServiceRegistry.Notes.TryParseAction,
                    (a, batch) => ServiceRegistry.Notes.DispatchToCore(_notesCommands, a, batch, request.Args)),
                "layout" => DispatchSimple<LayoutAction>(action, request,
                    ServiceRegistry.Layout.TryParseAction,
                    (a, batch) => ServiceRegistry.Layout.DispatchToCore(_layoutCommands, a, batch, request.Args)),
                "master" => DispatchSimple<MasterAction>(action, request,
                    ServiceRegistry.Master.TryParseAction,
                    (a, batch) => ServiceRegistry.Master.DispatchToCore(_masterCommands, a, batch, request.Args)),
                "animation" => DispatchSimple<AnimationAction>(action, request,
                    ServiceRegistry.Animation.TryParseAction,
                    (a, batch) => ServiceRegistry.Animation.DispatchToCore(_animationCommands, a, batch, request.Args)),
                "smartart" => DispatchSimple<SmartArtAction>(action, request,
                    ServiceRegistry.SmartArt.TryParseAction,
                    (a, batch) => ServiceRegistry.SmartArt.DispatchToCore(_smartArtCommands, a, batch, request.Args)),
                _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
            };

            return Task.FromResult(AttachRequestContext(request, response));
        }
        catch (Exception ex)
        {
            return Task.FromResult(CreateErrorResponse(ex, request.Command, request.SessionId));
        }
    }

    // === SERVICE COMMANDS ===

    private ServiceResponse HandleServiceCommand(string action)
    {
        return action switch
        {
            "ping" => new ServiceResponse { Success = true },
            "shutdown" => HandleShutdown(),
            "status" => HandleStatus(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown service action: {action}" }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        RequestShutdownAfterResponse();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = SessionCount,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            "apply-template" => HandleSessionApplyTemplate(request),
            "get-theme-name" => HandleSessionGetThemeName(request),
            "set-document-property" => HandleSessionSetDocumentProperty(request),
            "get-document-property" => HandleSessionGetDocumentProperty(request),
            "set-custom-property" => HandleSessionSetCustomProperty(request),
            "get-custom-property" => HandleSessionGetCustomProperty(request),
            "remove-custom-property" => HandleSessionRemoveCustomProperty(request),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var fullPath = Path.GetFullPath(args.FilePath);
        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use 'session open' to open an existing presentation."
            };
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Invalid file extension '{extension}'. session create supports .pptx and .pptm only."
            };
        }

        try
        {
            var sessionId = _sessions.Create(fullPath);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = fullPath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private ServiceResponse HandleSessionOpen(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        try
        {
            var sessionId = _sessions.Open(args.FilePath);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = args.FilePath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);

        if (args.Save && _sessions.TryGet(request.SessionId, out var batchToSave))
        {
            try
            {
                batchToSave.Save();
            }
            catch (Exception ex)
            {
                return CreateErrorResponse(ex, request.Command, request.SessionId);
            }
        }

        // Close() removes the session immediately and disposes the batch (and PowerPoint) on a
        // background task — mirrors PresentationSessionRegistry's MCP-server usage, see its
        // remarks for why this must not block the caller.
        var closed = _sessions.Close(request.SessionId);

        return new ServiceResponse { Success = true, Result = closed
            ? null
            : JsonSerializer.Serialize(new { success = true, sessionId = request.SessionId, message = "Session already closed." }, ServiceProtocol.JsonOptions) };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        if (!_sessions.TryGet(request.SessionId, out var batch))
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        try
        {
            batch.Save();
            return new ServiceResponse { Success = true };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    private ServiceResponse HandleSessionList()
    {
        var sessions = _sessions.List()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.PresentationPath,
                isPowerPointAlive = s.IsPowerPointProcessAlive
            })
            .ToList();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count }, ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse HandleSessionApplyTemplate(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionApplyTemplateArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.TemplatePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "templatePath is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.ApplyTemplate(batch!, args.TemplatePath), request);
    }

    private ServiceResponse HandleSessionGetThemeName(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        return WrapPresentationResult(() => _presentationCommands.GetThemeName(batch!), request);
    }

    private ServiceResponse HandleSessionSetDocumentProperty(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionPropertyArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.PropertyName))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "propertyName is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.SetDocumentProperty(batch!, args.PropertyName, args.Value ?? string.Empty), request);
    }

    private ServiceResponse HandleSessionGetDocumentProperty(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionPropertyArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.PropertyName))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "propertyName is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.GetDocumentProperty(batch!, args.PropertyName), request);
    }

    private ServiceResponse HandleSessionSetCustomProperty(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionPropertyArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.PropertyName))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "propertyName is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.SetCustomProperty(batch!, args.PropertyName, args.Value ?? string.Empty), request);
    }

    private ServiceResponse HandleSessionGetCustomProperty(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionPropertyArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.PropertyName))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "propertyName is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.GetCustomProperty(batch!, args.PropertyName), request);
    }

    private ServiceResponse HandleSessionRemoveCustomProperty(ServiceRequest request)
    {
        if (!TryGetBatch(request, out var batch, out var error))
        {
            return error!;
        }

        var args = ServiceRegistry.DeserializeArgs<SessionPropertyArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.PropertyName))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "propertyName is required" };
        }

        return WrapPresentationResult(() => _presentationCommands.RemoveCustomProperty(batch!, args.PropertyName), request);
    }

    /// <summary>
    /// Resolves <paramref name="request"/>'s sessionId to a live batch, mirroring the
    /// validation used by <see cref="DispatchSimple{TAction}"/> for the fully-generated
    /// domains. Used by the hand-written presentation actions (apply-template, theme, and
    /// document/custom property operations) that live under the "session" category.
    /// </summary>
    private bool TryGetBatch(ServiceRequest request, out IPresentationBatch? batch, out ServiceResponse? error)
    {
        batch = null;
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            error = new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
            return false;
        }

        if (!_sessions.TryGet(request.SessionId, out batch))
        {
            error = new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
            return false;
        }

        error = null;
        return true;
    }

    /// <summary>
    /// Invokes a hand-written <see cref="IPresentationCommands"/> operation and serializes its
    /// <see cref="PresentationOperationResult"/> the same way as the MCP <c>presentation</c>
    /// tool's <c>SerializeResult</c>, so CLI and MCP surfaces return identically-shaped JSON.
    /// </summary>
    private static ServiceResponse WrapPresentationResult(Func<PresentationOperationResult> invoke, ServiceRequest request)
    {
        try
        {
            var result = invoke();
            var json = JsonSerializer.Serialize(new
            {
                success = result.Success,
                errorMessage = result.ErrorMessage,
                presentationPath = result.PresentationPath,
                themeName = result.ThemeName,
                propertyName = result.PropertyName,
                propertyValue = result.PropertyValue
            }, ServiceProtocol.JsonOptions);

            return new ServiceResponse { Success = result.Success, Result = json, ErrorMessage = result.Success ? null : result.ErrorMessage };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    // === GENERATED DISPATCH ===

    private delegate bool TryParseDelegate<TAction>(string action, out TAction result);

    private static ServiceResponse WrapResult(string? dispatchResult)
    {
        return dispatchResult == null
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = true, Result = dispatchResult };
    }

    /// <summary>
    /// Resolves the session for <paramref name="request"/>, then dispatches
    /// <paramref name="actionString"/> to Core via the generated
    /// <c>ServiceRegistry.{Category}.DispatchToCore</c> method.
    /// </summary>
    private ServiceResponse DispatchSimple<TAction>(
        string actionString, ServiceRequest request,
        TryParseDelegate<TAction> tryParse,
        Func<TAction, IPresentationBatch, string?> dispatch) where TAction : struct
    {
        if (!tryParse(actionString, out var action))
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };
        }

        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        if (!_sessions.TryGet(request.SessionId, out var batch))
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        try
        {
            return WrapResult(dispatch(action, batch));
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    private static ServiceResponse AttachRequestContext(ServiceRequest request, ServiceResponse response)
    {
        if (response.Command != null && response.SessionId != null)
        {
            return response;
        }

        return new ServiceResponse
        {
            Success = response.Success,
            Command = response.Command ?? request.Command,
            SessionId = response.SessionId ?? request.SessionId,
            ErrorMessage = response.ErrorMessage,
            ErrorCategory = response.ErrorCategory,
            ExceptionType = response.ExceptionType,
            HResult = response.HResult,
            InnerError = response.InnerError,
            Result = response.Result
        };
    }

    private static ServiceResponse CreateErrorResponse(Exception ex, string? command = null, string? sessionId = null)
    {
        var exceptionType = ex.GetType().Name;
        string? hresult = ex is COMException comEx ? $"0x{comEx.HResult:X8}" : null;
        string? innerError = null;
        var errorCategory = ex switch
        {
            TimeoutException => "Timeout",
            ArgumentException => "InvalidInput",
            COMException => "ComInterop",
            _ => null
        };

        if (ex.InnerException != null)
        {
            innerError = ex.InnerException.Message;
            if (ex.InnerException is COMException innerComEx)
            {
                innerError += $" [COM: 0x{innerComEx.HResult:X8}]";
            }
        }

        return new ServiceResponse
        {
            Success = false,
            Command = command,
            SessionId = sessionId,
            ErrorCategory = errorCategory,
            ErrorMessage = ex.Message,
            ExceptionType = exceptionType,
            HResult = hresult,
            InnerError = innerError
        };
    }

    /// <summary>Disposes every open presentation session, closing PowerPoint for each.</summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _shutdownCts.Dispose();
        _sessions.DisposeAll();
    }
}

/// <summary>Args for "session.create"/"session.open".</summary>
public sealed class SessionOpenArgs
{
    /// <summary>Full path to the presentation file to create or open.</summary>
    public string? FilePath { get; set; }
}

/// <summary>Args for "session.close".</summary>
public sealed class SessionCloseArgs
{
    /// <summary>Whether to save the presentation before closing it.</summary>
    public bool Save { get; set; }
}

/// <summary>Args for "session.apply-template".</summary>
public sealed class SessionApplyTemplateArgs
{
    /// <summary>Full path to a .potx/.potm/.pot template file (or a .pptx/.pptm presentation used as a template source).</summary>
    public string? TemplatePath { get; set; }
}

/// <summary>Args for "session.set/get-document-property" and "session.set/get/remove-custom-property".</summary>
public sealed class SessionPropertyArgs
{
    /// <summary>Document property name (built-in or custom, depending on the action).</summary>
    public string? PropertyName { get; set; }

    /// <summary>The new property value. Used for set-document-property and set-custom-property.</summary>
    public string? Value { get; set; }
}
