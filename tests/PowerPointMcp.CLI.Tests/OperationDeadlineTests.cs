using Sbroenne.PowerPointMcp.CLI.Infrastructure;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class OperationDeadlineTests
{
    [Fact]
    public async Task Remaining_UsesOneMonotonicBudget()
    {
        var deadline = OperationDeadline.Start(TimeSpan.FromMilliseconds(100));
        var initial = deadline.Remaining;

        await Task.Delay(30);

        Assert.True(deadline.Remaining < initial);
        Assert.True(deadline.Cap(TimeSpan.FromSeconds(1)) <= deadline.Remaining);
    }
}
