using System.Runtime.InteropServices;
using Polly;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Pre-configured resilience pipelines for PowerPoint COM interop shutdown operations.
/// </summary>
/// <remarks>
/// Ported from mcp-server-excel's <c>ExcelMcp.ComInterop.Session.ResiliencePipelines</c>. Excel's
/// original also has Data Model and session-creation retry pipelines — those aren't needed here
/// (PowerPoint has no Data Model, and session creation doesn't retry today), so only the
/// Quit() pipeline is ported. Add more pipelines here if/when PowerPoint needs them.
/// Constant names use PascalCase (not Excel's SCREAMING_SNAKE_CASE) to satisfy this repo's
/// CA1707 analyzer rule, which — unlike mcp-server-excel's .editorconfig — is not suppressed.
/// </remarks>
public static class ResiliencePipelines
{
    #region COM HResult Constants

    /// <summary>RPC_E_SERVERCALL_RETRYLATER - COM server is busy, retry later.</summary>
    public const int RpcServerCallRetryLater = unchecked((int)0x8001010A);

    /// <summary>RPC_E_CALL_REJECTED - COM call was rejected (busy).</summary>
    public const int RpcCallRejected = unchecked((int)0x80010001);

    /// <summary>
    /// RPC_E_CALL_FAILED - RPC connection failed. PowerPoint is unreachable.
    /// FATAL - do not retry; the shutdown service falls through to process-exit polling
    /// and force-termination instead.
    /// </summary>
    public const int RpcCallFailed = unchecked((int)0x800706BE);

    /// <summary>
    /// RPC_S_SERVER_UNAVAILABLE - the RPC server is unavailable. The PowerPoint process has
    /// already died. FATAL - do not retry.
    /// </summary>
    public const int RpcServerUnavailable = unchecked((int)0x800706BA);

    /// <summary>
    /// RPC_E_DISCONNECTED - the COM proxy has disconnected from PowerPoint.
    /// FATAL - do not retry; proceed with cleanup.
    /// </summary>
    public const int RpcDisconnected = unchecked((int)0x80010108);

    #endregion

    /// <summary>
    /// Creates a retry pipeline for <c>Application.Quit()</c>: exponential backoff with jitter,
    /// retrying only transient COM-busy conditions (<see cref="RpcServerCallRetryLater"/>,
    /// <see cref="RpcCallRejected"/>). Fatal RPC failures (process already gone/disconnected)
    /// are deliberately excluded from retry so the caller moves straight to process-exit polling
    /// instead of retrying against a dead process.
    /// </summary>
    public static ResiliencePipeline CreatePowerPointQuitPipeline() =>
        new ResiliencePipelineBuilder()
            .AddRetry(new Polly.Retry.RetryStrategyOptions
            {
                MaxRetryAttempts = 6,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(500),
                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult != RpcCallFailed &&
                    ex.HResult != RpcServerUnavailable &&
                    ex.HResult != RpcDisconnected &&
                    (ex.HResult == RpcServerCallRetryLater ||
                     ex.HResult == RpcCallRejected)),
                OnRetry = static _ => ValueTask.CompletedTask
            })
            .Build();
}

