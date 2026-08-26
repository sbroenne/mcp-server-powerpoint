using System.Diagnostics;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

internal sealed class OperationDeadline
{
    private readonly long _startedAt = Stopwatch.GetTimestamp();
    private readonly TimeSpan _timeout;

    private OperationDeadline(TimeSpan timeout)
    {
        _timeout = timeout;
    }

    internal static OperationDeadline Start(TimeSpan timeout)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(timeout, TimeSpan.Zero);
        return new OperationDeadline(timeout);
    }

    internal TimeSpan Remaining
    {
        get
        {
            var remaining = _timeout - Stopwatch.GetElapsedTime(_startedAt);
            return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
        }
    }

    internal bool IsExpired => Remaining == TimeSpan.Zero;

    internal TimeSpan Cap(TimeSpan maximum)
    {
        var remaining = Remaining;
        var capped = remaining < maximum ? remaining : maximum;
        return capped > TimeSpan.FromMilliseconds(1)
            ? capped - TimeSpan.FromMilliseconds(1)
            : TimeSpan.Zero;
    }

    internal CancellationTokenSource CreateCancellationTokenSource() =>
        new(Remaining);
}
