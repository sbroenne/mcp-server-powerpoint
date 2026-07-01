using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Sbroenne.PowerPointMcp.ComInterop;

/// <summary>
/// OLE Message Filter for handling PowerPoint COM busy/retry scenarios.
/// Automatically retries when PowerPoint returns RPC_E_SERVERCALL_RETRYLATER.
/// </summary>
/// <remarks>
/// This filter intercepts COM calls to PowerPoint and handles transient "server busy" conditions.
/// When PowerPoint is temporarily busy (e.g., showing a dialog), the filter automatically retries
/// after a short delay rather than throwing an exception.
///
/// Register once per STA thread via Register(), revoke on thread shutdown via Revoke().
///
/// Ported from mcp-server-excel's OleMessageFilter, which is Office-COM-generic (not
/// Excel-specific). See that project's COM-fix-patterns history for why MessagePending
/// must return WAITDEFPROCESS (not WAITNOPROCESS) during long operations, and why
/// RetryRejectedCall must retry SERVERCALL_REJECTED for up to 2 minutes (enterprise
/// auth/sign-in dialogs cause repeated rejections) — both apply identically to PowerPoint.
/// </remarks>
[GeneratedComClass]
public sealed partial class OleMessageFilter : IOleMessageFilter
{
    private static readonly StrategyBasedComWrappers s_comWrappers = new();

    [ThreadStatic]
    private static nint _oldFilterPtr;

    [ThreadStatic]
    private static bool _isRegistered;

    /// <summary>
    /// When true, the filter is in a long-running COM operation (e.g., a slow export/render).
    /// MessagePending returns WAITDEFPROCESS to dispatch to HandleInComingCall, which rejects
    /// with SERVERCALL_RETRYLATER to trigger the caller's RetryRejectedCall backoff.
    /// </summary>
    [ThreadStatic]
    private static volatile bool _isInLongOperation;

    [ThreadStatic]
    private static long _messagePendingCount;

    [ThreadStatic]
    private static long _handleInComingCallRejections;

    [ThreadStatic]
    private static long _longOperationStartTimestamp;

    /// <summary>
    /// CancellationToken associated with the current outgoing COM call on this STA thread.
    /// When cancelled, <c>IMessageFilter.MessagePending</c> returns PENDINGMSG_CANCELCALL (0) to abort
    /// the pending outgoing call so the STA thread is not orphaned.
    /// </summary>
    [ThreadStatic]
    private static CancellationToken _pendingCancellationToken;

    /// <summary>
    /// Registers the OLE message filter for the current STA thread.
    /// Should be called once per STA thread before making COM calls.
    /// </summary>
    /// <exception cref="InvalidOperationException">Filter already registered on this thread, or registration failed</exception>
    public static void Register()
    {
        if (_isRegistered)
        {
            throw new InvalidOperationException("OLE message filter is already registered on this thread.");
        }

        var newFilter = new OleMessageFilter();
        nint newFilterPtr = s_comWrappers.GetOrCreateComInterfaceForObject(newFilter, CreateComInterfaceFlags.None);

        int result = CoRegisterMessageFilter(newFilterPtr, out _oldFilterPtr);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to register OLE message filter. HRESULT: 0x{result:X8}");
        }

        _isRegistered = true;
    }

    /// <summary>
    /// Revokes the OLE message filter and restores the previous filter.
    /// Safe to call even if Register() was not called - it will simply return.
    /// </summary>
    public static void Revoke()
    {
        if (!_isRegistered)
        {
            return;
        }

        int result = CoRegisterMessageFilter(_oldFilterPtr, out _);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to revoke OLE message filter. HRESULT: 0x{result:X8}");
        }

        _oldFilterPtr = 0;
        _isRegistered = false;
    }

    /// <summary>Gets whether the OLE message filter is registered on the current thread.</summary>
    public static bool IsRegistered => _isRegistered;

    /// <summary>Gets whether the filter is currently in a long operation on this thread.</summary>
    public static bool IsInLongOperation => _isInLongOperation;

    /// <summary>
    /// Marks the beginning of a long-running COM operation (e.g., slide export/render).
    /// </summary>
    public static void EnterLongOperation()
    {
        _messagePendingCount = 0;
        _handleInComingCallRejections = 0;
        _longOperationStartTimestamp = System.Diagnostics.Stopwatch.GetTimestamp();
        _isInLongOperation = true;
    }

    /// <summary>
    /// Marks the end of a long-running COM operation and returns diagnostic counters.
    /// </summary>
    public static (long MessagePendingCalls, long IncomingCallRejections, double ElapsedMs) ExitLongOperation()
    {
        _isInLongOperation = false;
        var elapsed = System.Diagnostics.Stopwatch.GetElapsedTime(_longOperationStartTimestamp);
        return (_messagePendingCount, _handleInComingCallRejections, elapsed.TotalMilliseconds);
    }

    /// <summary>
    /// Associates a <see cref="CancellationToken"/> with the current STA thread's outgoing COM call.
    /// </summary>
    public static void SetPendingCancellationToken(CancellationToken token)
    {
        _pendingCancellationToken = token;
    }

    /// <summary>
    /// Clears the pending cancellation token after the COM call completes or is cancelled.
    /// Must be called in a finally block after every <see cref="SetPendingCancellationToken"/> call.
    /// </summary>
    public static void ClearPendingCancellationToken()
    {
        _pendingCancellationToken = default;
    }

    /// <summary>
    /// Handles incoming COM calls. During long operations, rejects with SERVERCALL_RETRYLATER
    /// to trigger the caller's RetryRejectedCall backoff, preventing CPU spin from re-entrant dispatch.
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, nint htaskCaller, int dwTickCount, nint lpInterfaceInfo)
    {
        if (_isInLongOperation)
        {
            Interlocked.Increment(ref _handleInComingCallRejections);
            return 2; // SERVERCALL_RETRYLATER
        }

        return 0; // SERVERCALL_ISHANDLED
    }

    /// <summary>
    /// Handles rejected COM calls from PowerPoint.
    /// Implements automatic retry logic with exponential backoff for busy/unavailable conditions.
    /// </summary>
    int IOleMessageFilter.RetryRejectedCall(nint htaskCallee, int dwTickCount, int dwRejectType)
    {
        const int SERVERCALL_REJECTED = 1;
        const int SERVERCALL_RETRYLATER = 2;
        const int RETRY_LATER_TIMEOUT_MS = 30000;
        const int REJECTED_TIMEOUT_MS = 120000; // 2 minutes for auth/sign-in dialogs

        if (dwRejectType == SERVERCALL_RETRYLATER)
        {
            if (dwTickCount >= RETRY_LATER_TIMEOUT_MS)
            {
                return -1;
            }

            return dwTickCount switch
            {
                < 1000 => 100,
                < 5000 => 200,
                < 15000 => 500,
                _ => 1000
            };
        }

        if (dwRejectType == SERVERCALL_REJECTED)
        {
            if (dwTickCount >= REJECTED_TIMEOUT_MS)
            {
                return -1;
            }

            return 1000;
        }

        return -1;
    }

    /// <summary>
    /// Handles pending message during a COM call. During long operations, dispatches to
    /// HandleInComingCall (which rejects). During normal operations, queues messages without
    /// dispatching, to avoid re-entrant COM execution on the STA thread.
    /// </summary>
    int IOleMessageFilter.MessagePending(nint htaskCallee, int dwTickCount, int dwPendingType)
    {
        Interlocked.Increment(ref _messagePendingCount);

        if (_pendingCancellationToken.IsCancellationRequested)
        {
            return 0; // PENDINGMSG_CANCELCALL
        }

        if (_isInLongOperation)
        {
            return 2; // PENDINGMSG_WAITDEFPROCESS
        }

        return 1; // PENDINGMSG_WAITNOPROCESS
    }

    /// <summary>Registers or revokes a message filter for the current apartment.</summary>
    [LibraryImport("Ole32.dll")]
    private static partial int CoRegisterMessageFilter(
        nint lpMessageFilter,
        out nint lplpMessageFilter);
}
