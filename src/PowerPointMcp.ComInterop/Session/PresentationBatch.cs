using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Win32;
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
/// multi-presentation batches, IRM/AIP detection, and macro-security handling for .pptm.
/// Resilient close/quit retry + process-exit polling IS implemented — see
/// <see cref="PresentationShutdownService"/>.
/// </para>
/// </remarks>
internal sealed class PresentationBatch : IPresentationBatch
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool PeekMessage(out MSG msg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG msg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG msg);

    [DllImport("user32.dll")]
    private static extern uint MsgWaitForMultipleObjectsEx(uint nCount, IntPtr[] pHandles, uint dwMilliseconds, uint dwWakeMask, uint dwFlags);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public int pt_x;
        public int pt_y;
    }

    private string _presentationPath;
    private readonly bool _showPowerPoint;
    private readonly bool _createNewFile;
    private readonly TimeSpan _operationTimeout;
    private readonly ILogger<PresentationBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private readonly AutoResetEvent _workSignal = new(false);
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

    /// <summary>
    /// Sets the (undocumented, per-user) PowerPoint Designer preference that disables the
    /// "Automatically show me design ideas" feature.
    /// </summary>
    /// <remarks>
    /// <para><b>Discovered via real integration testing, not assumed:</b> calling
    /// <c>Presentation.ApplyTemplate</c> on a slide that already has shape/text content
    /// triggers PowerPoint's Designer (Design Ideas) feature, which opens a blocking, title-less
    /// modal dialog of window class <c>NUIDialog</c>. With no interactive user present to dismiss
    /// it, the call hangs forever until our own <see cref="ComInteropConstants.DefaultOperationTimeout"/>
    /// force-kills the process.</para>
    /// <para>
    /// There is no supported COM/VBA property to toggle this (confirmed: not exposed via
    /// <c>Application</c> or <c>Presentation</c>). The supported Group Policy key
    /// (<c>HKCU\Software\Policies\Microsoft\Office\16.0\PowerPoint\Options\DisableDesigner</c>) is
    /// commonly locked down by MDM/policy and not writable from a normal user context. The key
    /// set here (<c>HKCU\Software\Microsoft\Office\16.0\PowerPoint\Designer\DisableAutomaticDesigner</c>)
    /// is the plain per-user preference behind "File &gt; Options &gt; General &gt; PowerPoint
    /// Designer &gt; Automatically show me design ideas" — undocumented/unsupported by Microsoft,
    /// but empirically verified (via direct COM repro) to eliminate the hang while leaving slide
    /// content and theme application fully correct.
    /// </para>
    /// <para>
    /// This is a deliberate, best-effort side effect: it changes the *user's* PowerPoint Designer
    /// preference machine-wide (HKCU), not just for this automation session, so that Designer
    /// won't unexpectedly hang ANY future automated session either. Failure to set it (e.g.
    /// read-only registry, non-Windows test harness) is swallowed — automation continues without
    /// the mitigation rather than failing startup.
    /// </para>
    /// </remarks>
    private void DisableDesignerAutomationHang()
    {
        try
        {
            using RegistryKey key = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Office\16.0\PowerPoint\Designer");
            key.SetValue("DisableAutomaticDesigner", 1, RegistryValueKind.DWord);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Could not set the PowerPoint Designer preference registry key to prevent the " +
                "known ApplyTemplate/Designer hang. Automation will continue, but ApplyTemplate " +
                "calls on slides with existing content may hang until the operation timeout.");
        }
    }

    private void RunStaThread(TaskCompletionSource started)
    {
        PowerPoint.Application? startupApp = null;
        PowerPoint.Presentation? startupPresentation = null;
        try
        {
            OleMessageFilter.Register();
            DisableDesignerAutomationHang();

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

            dynamic dynApp = tempApp;
            // NOTE (discovered via real integration test, not assumed): PowerPoint's automation
            // COM object does NOT allow hiding its application window — setting
            // Application.Visible = False throws COMException "Hiding the application window is
            // not allowed." (unlike Excel, which does allow it). So PowerPoint is unconditionally
            // shown; _showPowerPoint is retained for API parity with mcp-server-excel's show
            // option and any future distinction we might add (e.g. window position), but does not
            // currently toggle Application.Visible.
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

            string fullPath = Path.GetFullPath(_presentationPath);
            PowerPoint.Presentation presentation = OpenOrCreatePresentationCom(dynApp, fullPath, _createNewFile);

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

            PresentationShutdownService.CloseAndQuit(cleanupPresentation, cleanupApp, _powerPointProcessId, _logger, _presentationPath);

            _presentation = null;
            _app = null;
            _context = null;

            try { OleMessageFilter.Revoke(); }
            catch (Exception ex) { _logger.LogWarning(ex, "OleMessageFilter.Revoke() failed during STA cleanup"); }
        }
    }

    /// <summary>
    /// Creates a new blank presentation (createNewFile=true) or opens an existing one
    /// (createNewFile=false) against the given live PowerPoint.Application, using the same
    /// COM parameters/rationale as the initial STA-thread startup path. Extracted so the
    /// same logic can be reused by <see cref="ReopenPresentation"/> to swap presentations
    /// within an already-running Application instance (e.g. shared test fixtures) without
    /// re-launching PowerPoint.
    /// </summary>
    private static PowerPoint.Presentation OpenOrCreatePresentationCom(dynamic dynApp, string fullPath, bool createNewFile)
    {
        const int msoTrue = -1;
        const int msoFalse = 0;

        PowerPoint.Presentation presentation;

        if (createNewFile)
        {
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'.");
            }

            // See RunStaThread's original NOTE: WithWindow must be msoTrue (chart in-place
            // activation needs a real document window), independent of Application.Visible.
            presentation = (PowerPoint.Presentation)dynApp.Presentations.Add(msoTrue);

            // Presentations.Add() creates a presentation with ZERO slides — add one blank
            // slide so "create a new presentation" is immediately useful.
            presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            int formatCode = string.Equals(Path.GetExtension(fullPath), ".pptm", StringComparison.OrdinalIgnoreCase)
                ? ComInteropConstants.PpSaveAsOpenXmlPresentationMacroEnabled
                : ComInteropConstants.PpSaveAsOpenXmlPresentation;
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
                msoFalse, // ReadOnly = false — we need to be able to Save()
                msoFalse, // Untitled = false — see RunStaThread's original NOTE.
                msoTrue); // WithWindow = true — see RunStaThread's original NOTE.
        }

        return presentation;
    }

    private void ProcessWorkQueue()
    {
        // NOTE (discovered via real integration test, not assumed): PowerPoint's Chart feature
        // (Shapes.AddChart2 / Chart.ChartData) embeds a separate, out-of-process Excel instance
        // for chart data editing. Cross-process COM calls into/out of that embedded Excel
        // instance require this STA thread to actually pump Windows messages (SendMessage-based
        // RPC callbacks rely on GetMessage/DispatchMessage) — a plain async Channel.Reader wait
        // (no message pump at all) causes those calls to fail with a generic
        // COMException(0x80004005) "Unexpected HRESULT". OleMessageFilter alone does not
        // substitute for this: it governs COM's busy/reject retry behavior, not Win32 message
        // dispatch. Fixed by replacing the blocking async wait with a native message pump
        // (PeekMessage/TranslateMessage/DispatchMessage) interleaved with draining the work
        // queue, waking on either new work (_workSignal) or new Windows messages
        // (MsgWaitForMultipleObjectsEx with QS_ALLINPUT).
        const uint PM_REMOVE = 0x0001;
        const uint QS_ALLINPUT = 0x04FF;
        const uint MWMO_INPUTAVAILABLE = 0x0004;
        const uint pollTimeoutMs = 50;

        IntPtr[] waitHandles =
        [
            _workSignal.SafeWaitHandle.DangerousGetHandle(),
            _shutdownCts.Token.WaitHandle.SafeWaitHandle.DangerousGetHandle()
        ];

        while (!_shutdownCts.IsCancellationRequested)
        {
            while (_workQueue.Reader.TryRead(out var work))
            {
                try { work().GetAwaiter().GetResult(); }
                catch (Exception) { /* already captured in the caller's TaskCompletionSource */ }
            }

            while (PeekMessage(out MSG msg, IntPtr.Zero, 0, 0, PM_REMOVE))
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }

            _ = MsgWaitForMultipleObjectsEx((uint)waitHandles.Length, waitHandles, pollTimeoutMs, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        }

        // Drain any remaining queued work after shutdown was signaled.
        while (_workQueue.Reader.TryRead(out var remainingWork))
        {
            try { remainingWork().GetAwaiter().GetResult(); }
            catch (Exception) { /* already captured */ }
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
            _workSignal.Set();
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

    /// <summary>
    /// Closes the currently-open presentation (WITHOUT quitting the Application) and then
    /// creates a new blank presentation or opens an existing one, reusing the same live
    /// PowerPoint.Application instance and STA thread. This lets callers (currently: shared
    /// per-test-class fixtures in Core.Tests) pay the PowerPoint launch/teardown cost once per
    /// Application instance while still giving every test its own fresh presentation file for
    /// isolation. Not exposed on <see cref="IPresentationBatch"/> — this is an internal
    /// test-performance affordance, not a general multi-presentation feature (see the "Scope
    /// note" on <see cref="IPresentationBatch"/> re: single-presentation-per-batch as the
    /// production contract).
    /// </summary>
    internal void ReopenPresentation(string filePath, bool createNewFile, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("A file path is required", nameof(filePath));

        string fullPath = Path.GetFullPath(filePath);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (createNewFile)
        {
            if (extension is not (".pptx" or ".pptm"))
                throw new ArgumentException($"Invalid file extension '{extension}'. New presentations must be .pptx or .pptm.");
        }
        else if (extension is not (".pptx" or ".pptm" or ".ppt"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only PowerPoint files (.pptx, .pptm, .ppt) are supported.");
        }

        Execute((ctx, ct) =>
        {
            var oldPresentation = _presentation;
            if (oldPresentation != null)
            {
                try { oldPresentation.Close(); }
                catch { /* best effort - proceed with opening the next presentation regardless */ }
                ComUtilities.Release(ref oldPresentation);
            }

            dynamic dynApp = _app!;
            PowerPoint.Presentation newPresentation = OpenOrCreatePresentationCom(dynApp, fullPath, createNewFile);
            _presentation = newPresentation;
            _presentationPath = fullPath;
            _context = new PresentationContext(fullPath, _app!, newPresentation);
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

        bool staThreadExited = _staThread.Join(ComInteropConstants.StaThreadJoinTimeout);

        // Ultimate backstop: StaThreadJoinTimeout is sized to comfortably cover
        // PresentationShutdownService's full worst-case duration (Close/Quit retries + the
        // process-exit grace period + its own force-kill). If the STA thread STILL hasn't
        // finished even after that — e.g. Quit() itself is blocked on a modal dialog that
        // outlasts every other safety net — force-kill here so Dispose() never returns having
        // silently leaked the process.
        if (!staThreadExited && _powerPointProcessId.HasValue)
        {
            TryKillProcess(_powerPointProcessId.Value);
        }

        _shutdownCts.Dispose();
        _workSignal.Dispose();
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
