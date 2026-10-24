using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "PresentationBatch")]
public sealed class PresentationShutdownTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SessionEnd_PreservesUnsavedWorkOnlyDuringHostShutdown(bool explicitClose)
    {
        string path = Path.Combine(Path.GetTempPath(), $"pptmcp-unsaved-{Guid.NewGuid():N}.pptx");
        string observerPath = Path.Combine(Path.GetTempPath(), $"pptmcp-observer-{Guid.NewGuid():N}.pptx");
        using var observer = PresentationSession.CreateNew(observerPath);
        using var registry = new PresentationSessionRegistry();
        try
        {
            string sessionId = registry.Create(path);
            Assert.True(registry.TryGet(sessionId, out var batch));
            batch.Execute((context, token) =>
            {
                PowerPoint.Slides? slides = null;
                PowerPoint.Slide? slide = null;
                try
                {
                    slides = context.Presentation.Slides;
                    slide = slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                }
                finally
                {
                    ComUtilities.Release(ref slide);
                    ComUtilities.Release(ref slides);
                }
            });
            byte[] savedFile = File.ReadAllBytes(path);

            if (explicitClose)
            {
                batch.BeginRequest();
                Assert.True(registry.Close(sessionId));
                batch.CompleteRequest("completed");
            }
            registry.DisposeAll();

            bool retained = observer.Execute((context, token) =>
            {
                PowerPoint.Presentations? presentations = null;
                try
                {
                    presentations = context.App.Presentations;
                    for (int i = 1; i <= presentations.Count; i++)
                    {
                        PowerPoint.Presentation? presentation = null;
                        try
                        {
                            presentation = presentations[i];
                            if (presentation.FullName == path)
                            {
                                context.App.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                                presentation.Close();
                                return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref presentation);
                        }
                    }
                    return false;
                }
                finally
                {
                    ComUtilities.Release(ref presentations);
                }
            });
            if (explicitClose)
                Assert.False(retained, "An explicit close must close the presentation even within its owning request.");
            else
                Assert.True(retained, "Host shutdown must leave unsaved work open in PowerPoint.");
            Assert.Equal(savedFile, File.ReadAllBytes(path));
        }
        finally
        {
            registry.Dispose();
            observer.Dispose();
            File.Delete(path);
            File.Delete(observerPath);
        }
    }

    [Fact]
    public void HostShutdown_RejectsLaterStartup()
    {
        using var registry = new PresentationSessionRegistry();
        registry.DisposeAll();
        Assert.Throws<ObjectDisposedException>(() => registry.Create("unused.pptx"));
    }
}
