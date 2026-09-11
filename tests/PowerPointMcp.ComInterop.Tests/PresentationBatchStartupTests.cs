using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.ComInterop.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "PresentationBatch")]
public sealed class PresentationBatchStartupTests
{
    [Fact]
    public void CreateNew_StartupTimeoutAfterApplicationCreation_DoesNotLeavePowerPointRunning()
    {
        PowerPointProcessIdentity? createdIdentity = null;
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-startup-{Guid.NewGuid():N}.pptx");

        PresentationBatch.AfterProcessIdentityCapturedHook = (identity, cancellationToken) =>
        {
            createdIdentity = identity;
            cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(30));
        };

        bool processWasAliveAfterFailure = false;
        try
        {
            _ = Assert.Throws<TimeoutException>(() =>
                PresentationSession.CreateNew(path, operationTimeout: TimeSpan.FromSeconds(2)));

            Assert.NotNull(createdIdentity);
            processWasAliveAfterFailure = OwnedProcessGuard.IsAlive(createdIdentity.Value);
        }
        finally
        {
            PresentationBatch.AfterProcessIdentityCapturedHook = null;

            if (createdIdentity is { } identity
                && OwnedProcessGuard.TryOpenMatchingProcess(identity, out var process))
            {
                using (process)
                {
                    if (process is not null)
                    {
                        process.Kill();
                        process.WaitForExit(5000);
                    }
                }
            }

            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
        }

        Assert.False(
            processWasAliveAfterFailure,
            "A PowerPoint process created during failed startup remained alive after the constructor returned.");
    }
}
