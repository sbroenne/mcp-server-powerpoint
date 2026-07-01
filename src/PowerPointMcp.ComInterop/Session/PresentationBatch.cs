using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Implementation of IPresentationBatch that manages a single PowerPoint instance on a
/// dedicated STA thread. Ensures proper COM interop with PowerPoint using STA apartment
/// state and OLE message filter.
/// </summary>
/// <remarks>
/// <para><b>CRITICAL: PowerPoint COM Threading Model</b> (same rules as Excel — see
/// mcp-server-excel's ExcelBatch for the original, more heavily hardened implementation)</para>
/// <list type="bullet">
/// <item>Each PresentationBatch runs on ONE dedicated STA thread</item>
/// <item>Operations are queued via Channel and executed SERIALLY (never in parallel)</item>
/// <item>This is a COM interop requirement, not an implementation choice</item>
/// </list>
/// <para>
/// <b>Scope note:</b> this is a first-pass port focused on proving the architecture
/// (open/create/close/save). It intentionally omits several hardening features present in
/// ExcelBatch that should be ported before this is considered production-ready:
/// multi-presentation batches, IRM/AIP detection, macro-security handling for .pptm,
/// force-kill-on-timeout diagnostics, and a full PresentationShutdownService with
/// exponential-backoff retry (a minimal inline close/quit is used here instead).
/// </para>
/// </remarks>
internal sealed class PresentationBatch : IPresentationBatch
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private readonly string _presentationPath;
    private readonly bool _showPowerPoint;
    private readonly bool _createNewFile;
    private readonly TimeSpan _operationTimeout;
    private readonly ILogger<PresentationBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed;
    private int? _powerPointProcessId;
    private bool _operationTimedOut;

    private PowerPoint.Application? _app;
    private PowerPoint.Presentation? _presentation;
    private PresentationContext? _context;

    public PresentationBatch(
        string presentationPath,
        bool createNewFile,
        ILogger<PresentationBatch>? logger = null,
        bool show = false,
        TimeSpan? operationTimeout = null)
    {
        _presentationPath = presentationPath ?? throw new ArgumentNullException(nameof(presentationPath));
        _createNewFile = createNewFile;
        _showPowerPoint = show;
        _operationTimeout = operationTimeout ?? ComInteropConstants.DefaultOperationTimeout;
        _logger = logger ?? NullLogger<PresentationBatch>.Instance;
        _shutdownCts = new CancellationTokenSource();

        _workQueue = Channel.CreateUnbounded<Func<Task>>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(() => RunStaThread(started))
        {
            IsBackground = true,
            Name = $"PresentationBatch-{Path.GetFileName(_presentationPath)}"
        };

        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        try
        {
            bool completedInTime = started.Task.Wait(_operationTimeout);
            if (!completedInTime)
            {
                _operationTimedOut = true;
                _workQueue.Writer.TryComplete();
                _shutdownCts.Cancel();
                throw new TimeoutException(
                    $"PowerPoint startup timed out after {_operationTimeout.TotalSeconds} seconds while opening " +
                    $"'{Path.GetFileName(_presentationPath)}'. The presentation may be blocked on an interactive " +
                    "dialog or an unresponsive open. Retry with a larger operationTimeout, or with show=true so " +
                    "PowerPoint is visible for prompts.");
            }

            started.Task.GetAwaiter().GetResult();
        }
        catch
        {
            if (_staThread.IsAlive)
            {
                if (!_staThread.Join(TimeSpan.FromSeconds(10)) && _powerPointProcessId.HasValue)
                {
                    TryKillProcess(_powerPointProcessId.Value);
                    _ = _staThread.Join(TimeSpan.FromSeconds(5));
                }
            }
            throw;
        }
    }

    private void RunStaThread(TaskCompletionSource started)
    {
        PowerPoint.Application? startupApp = null;
        PowerPoint.Presentation? startupPresentation = null;
        try
        {
            OleMessageFilter.Register();

            Type? appType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (appType == null)
            {
                throw new InvalidOperationException("Microsoft PowerPoint is not installed on this system.");
            }

            var tempApp = (PowerPoint.Application)Activator.CreateInstance(appType)!;
            startupApp = tempApp;

            // NOTE: PowerPoint.Application.Visible and Presentations.Add/Open all take
            // Microsoft.Office.Core.MsoTriState parameters. We deliberately do NOT reference
            // office.dll (Microsoft.Office.Core) — same rationale as mcp-server-excel's
            // AutomationSecurity trick: access these late-bound via dynamic (IDispatch), passing
            // the raw MsoTriState int values (msoTrue = -1, msoFalse = 0). This keeps the
            // assembly free of any office.dll runtime/compile dependency.
            const int msoTrue = -1;
            const int msoFalse = 0;

            dynamic dynApp = tempApp;
            try { dynApp.Visible = msoTrue; } catch { /* best effort */ }
            tempApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

            try
            {
                const int maxRetries = 3;
                const int retryDelayMs = 500;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    int hwnd = tempApp.HWND;
                    if (hwnd != 0)
                    {
                        _ = GetWindowThreadProcessId(new IntPtr(hwnd), out uint processId);
                        if (processId != 0)
                        {
                            _powerPointProcessId = (int)processId;
                            break;
                        }
                    }
                    if (attempt < maxRetries) Thread.Sleep(retryDelayMs);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to capture PowerPoint process ID. Force-kill will not be available.");
            }

            PowerPoint.Presentation presentation;
            string fullPath = Path.GetFullPath(_presentationPath);

            if (_createNewFile)
            {
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'.");
                }

                presentation = (PowerPoint.Presentation)dynApp.Presentations.Add(msoFalse);
                int formatCode = string.Equals(Path.GetExtension(fullPath), ".pptm", StringComparison.OrdinalIgnoreCase)
                    ? ComInteropConstants.PpSaveAsOpenXmlPresentationMacroEnabled
                    : ComInteropConstants.PpSaveAsOpenXmlPresentation;
                // Late-bound call: Presentation.SaveAs's third parameter (EmbedTrueTypeFonts) is
                // typed as Microsoft.Office.Core.MsoTriState. Calling it statically would force a
                // reference to office.dll just to satisfy the default-parameter type. Dynamic
                // dispatch avoids that entirely (mirrors the AutomationSecurity trick above).
                ((dynamic)presentation).SaveAs(fullPath, formatCode);
            }
            else
            {
                if (!File.Exists(fullPath))
                {
                    throw new FileNotFoundException($"PowerPoint file not found: {fullPath}.", fullPath);
                }

                presentation = (PowerPoint.Presentation)dynApp.Presentations.Open(
                    fullPath,
                    msoFalse,
                    msoTrue,
                    _showPowerPoint ? msoTrue : msoFalse);
            }

            startupPresentation = presentation;
            _app = tempApp;
            _presentation = presentation;
            _context = new PresentationContext(fullPath, tempApp, presentation);

            started.SetResult();

            ProcessWorkQueue();
        }
        catch (Exception ex)
        {
            started.TrySetException(ex);
        }
        finally
        {
            var cleanupApp = _app ?? startupApp;
            var cleanupPresentation = _presentation ?? startupPresentation;

            PresentationShutdownService.CloseAndQuit(cleanupPresentation, cleanupApp, _logger);

            _presentation = null;
            _app = null;
            _context = null;

            try { OleMessageFilter.Revoke(); }
            catch (Exception ex) { _logger.LogWarning(ex, "OleMessageFilter.Revoke() failed during STA cleanup"); }
        }
    }

    private void ProcessWorkQueue()
    {
        try
        {
            while (true)
            {
                ValueTask<bool> waitTask = _workQueue.Reader.WaitToReadAsync(_shutdownCts.Token);
                bool hasData = waitTask.IsCompleted ? waitTask.Result : waitTask.AsTask().GetAwaiter().GetResult();
                if (!hasData) break;

                while (_workQueue.Reader.TryRead(out var work))
                {
                    try { work().GetAwaiter().GetResult(); }
                    catch (Exception) { /* already captured in the caller's TaskCompletionSource */ }
                }
            }
        }
        catch (OperationCanceledException)
        {
            while (_workQueue.Reader.TryRead(out var remainingWork))
            {
                try { remainingWork().GetAwaiter().GetResult(); }
                catch (Exception) { /* already captured */ }
            }
        }
    }

    public string PresentationPath => _presentationPath;

    public bool HasTimedOutOperation => _operationTimedOut;

    public int? PowerPointProcessId => _powerPointProcessId;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsPowerPointProcessAlive()
    {
        if (!_powerPointProcessId.HasValue) return false;
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(_powerPointProcessId.Value);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    public void Execute(Action<PresentationContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        => Execute((ctx, ct) => { operation(ctx, ct); return 0; }, cancellationToken);

    public T Execute<T>(Func<PresentationContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(PresentationBatch));

        if (_operationTimedOut)
        {
            throw new TimeoutException(
                $"A previous operation timed out for '{Path.GetFileName(_presentationPath)}'. " +
                "Please close this session and create a new one.");
        }

        if (!IsPowerPointProcessAlive())
        {
            throw new InvalidOperationException(
                $"PowerPoint process is no longer running for '{Path.GetFileName(_presentationPath)}'. " +
                "It may have been closed manually or crashed. Please close this session and create a new one.");
        }

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        try
        {
            var writeTask = _workQueue.Writer.WriteAsync(() =>
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var result = operation(_context!, cancellationToken);
                    tcs.SetResult(result);
                }
                catch (OperationCanceledException oce)
                {
                    tcs.TrySetCanceled(oce.CancellationToken);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
                return Task.CompletedTask;
            }, cancellationToken);

            if (writeTask.IsCompleted) writeTask.GetAwaiter().GetResult();
            else writeTask.AsTask().GetAwaiter().GetResult();
        }
        catch (ChannelClosedException)
        {
            throw new ObjectDisposedException(nameof(PresentationBatch),
                $"Session for '{Path.GetFileName(_presentationPath)}' was disposed while submitting an operation.");
        }

        try
        {
            if (cancellationToken.CanBeCanceled)
            {
                return tcs.Task.WaitAsync(cancellationToken).GetAwaiter().GetResult();
            }

            using var timeoutCts = new CancellationTokenSource(_operationTimeout);
            return tcs.Task.WaitAsync(timeoutCts.Token).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            _operationTimedOut = true;
            throw new TimeoutException(
                $"PowerPoint operation timed out after {_operationTimeout.TotalSeconds} seconds for " +
                $"'{Path.GetFileName(_presentationPath)}'.");
        }
        catch (OperationCanceledException)
        {
            _operationTimedOut = true;
            throw;
        }
    }

    public void Save(CancellationToken cancellationToken = default)
    {
        Execute((ctx, ct) =>
        {
            ctx.Presentation.Save();
            return 0;
        }, cancellationToken);
    }

    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0) return;

        _shutdownCts.Cancel();
        _workQueue.Writer.Complete();

        if (_operationTimedOut && _powerPointProcessId.HasValue && _staThread.IsAlive)
        {
            TryKillProcess(_powerPointProcessId.Value);
        }

        _staThread.Join(ComInteropConstants.StaThreadJoinTimeout);
        _shutdownCts.Dispose();
    }

    private static void TryKillProcess(int processId)
    {
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(processId);
            if (!process.HasExited)
            {
                process.Kill();
                process.WaitForExit(5000);
            }
        }
        catch (ArgumentException) { /* already exited */ }
        catch (Exception) { /* best effort */ }
    }
}
