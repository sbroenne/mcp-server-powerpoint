using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.ComInterop.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "PresentationBatch")]
public sealed class PresentationBatchStartupTests
{
    [Fact]
    public void StartupFailure_DoesNotEndExistingPowerPoint()
    {
        string existingPath = Path.Combine(Path.GetTempPath(), $"pptmcp-startup-observer-{Guid.NewGuid():N}.pptx");
        string failedPath = Path.Combine(Path.GetTempPath(), $"pptmcp-startup-failed-{Guid.NewGuid():N}.pptx");
        using var observer = PresentationSession.CreateNew(existingPath);
        try
        {
            PresentationBatch.AfterProcessIdentityCapturedHook = (_, _) =>
                throw new InvalidOperationException("Simulated startup failure.");
            Assert.ThrowsAny<Exception>(() => PresentationSession.CreateNew(failedPath));

            Assert.True(observer.IsPowerPointProcessAlive(), "Failed startup must not terminate a shared PowerPoint application.");
            Assert.Contains(observer.PowerPointProcessIdentity!.Value,
                PresentationSessionRegistry.GetTrackedPowerPointProcesses());
            Assert.Equal(Path.GetFileName(existingPath), observer.Execute((context, token) => context.Presentation.Name));
        }
        finally
        {
            PresentationBatch.AfterProcessIdentityCapturedHook = null;
            observer.Dispose();
            File.Delete(existingPath);
            File.Delete(failedPath);
        }
    }

    [Fact]
    public void CreateNew_StartupTimeoutAfterApplicationCreation_ReleasesAutomationTracking()
    {
        PowerPointProcessIdentity? createdIdentity = null;
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-startup-{Guid.NewGuid():N}.pptx");

        PresentationBatch.AfterProcessIdentityCapturedHook = (identity, cancellationToken) =>
        {
            createdIdentity = identity;
            cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(30));
        };

        try
        {
            _ = Assert.Throws<TimeoutException>(() =>
                PresentationSession.CreateNew(path, operationTimeout: TimeSpan.FromSeconds(2)));

            Assert.NotNull(createdIdentity);
            Assert.DoesNotContain(createdIdentity.Value, PresentationSessionRegistry.GetTrackedPowerPointProcesses());
        }
        finally
        {
            PresentationBatch.AfterProcessIdentityCapturedHook = null;

            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
        }

    }
}
