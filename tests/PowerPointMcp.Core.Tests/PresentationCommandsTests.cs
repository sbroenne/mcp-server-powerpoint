using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests against a live PowerPoint COM instance. NO mocking — per Rule 30
/// (integration tests over unit tests), these require PowerPoint to be installed and drive
/// the actual Presentations.Add/Open/SaveAs/Close/Quit COM calls.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Presentation")]
public class PresentationCommandsTests
{
    private readonly PresentationCommands _commands = new();

    [Fact]
    public void SaveAs_Pptx_ReopensWithEditedContentAndUpdatesSessionPath()
    {
        string originalPath = CoreTestHelper.CreateUniqueTestFilePath();
        string targetPath = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            using (var batch = PresentationSession.CreateNew(originalPath))
            {
                batch.Execute((ctx, ct) =>
                {
                    AddBlankSlide(ctx);
                });

                var result = _commands.SaveAs(
                    batch,
                    targetPath,
                    PresentationSaveFormat.Pptx);

                Assert.True(result.Success);
                Assert.Null(result.ErrorMessage);
                Assert.Equal(Path.GetFullPath(targetPath), result.PresentationPath, ignoreCase: true);
                Assert.Equal(Path.GetFullPath(targetPath), batch.PresentationPath, ignoreCase: true);
                string contextPath = batch.Execute((ctx, ct) => ctx.PresentationPath);
                Assert.Equal(Path.GetFullPath(targetPath), contextPath, ignoreCase: true);
            }

            using var reopened = PresentationSession.BeginBatch(targetPath);
            int slideCount = reopened.Execute((ctx, ct) => GetSlideCount(ctx));
            Assert.Equal(2, slideCount);
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(targetPath);
        }
    }

    [Fact]
    public void SaveAs_ComFailureAfterValidation_PreservesBatchAndRegistryPath()
    {
        string originalPath = CoreTestHelper.CreateUniqueTestFilePath();
        string outputDirectory = Path.Combine(
            Path.GetTempPath(),
            "PowerPointMcpTests",
            $"save-as-failure-{Guid.NewGuid():N}");
        string targetPath = Path.Combine(outputDirectory, "failed-save.pptx");
        Directory.CreateDirectory(outputDirectory);
        var registry = new PresentationSessionRegistry();

        try
        {
            string sessionId = registry.Create(originalPath);
            Assert.True(registry.TryGet(sessionId, out var batch));
            batch.Save();

            bool executeStarted = false;
            var interceptingBatch = new BeforeExecuteBatch(batch, () =>
            {
                executeStarted = true;
                Directory.Delete(outputDirectory);
            });

            Assert.Throws<COMException>(() => _commands.SaveAs(
                interceptingBatch,
                targetPath,
                PresentationSaveFormat.Pptx));

            Assert.True(executeStarted);
            Assert.Equal(Path.GetFullPath(originalPath), batch.PresentationPath, ignoreCase: true);
            Assert.Equal(
                Path.GetFullPath(originalPath),
                batch.Execute((ctx, ct) => ctx.PresentationPath),
                ignoreCase: true);
            var session = Assert.Single(registry.List());
            Assert.Equal(sessionId, session.SessionId);
            Assert.Equal(
                Path.GetFullPath(originalPath),
                session.PresentationPath,
                ignoreCase: true);
            Assert.False(File.Exists(targetPath));
        }
        finally
        {
            registry.Dispose();
            File.Delete(originalPath);
            if (Directory.Exists(outputDirectory))
            {
                Directory.Delete(outputDirectory, recursive: true);
            }
        }
    }

    private static void AddBlankSlide(PresentationContext context)
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
    }

    private static int GetSlideCount(PresentationContext context)
    {
        PowerPoint.Slides? slides = null;
        try
        {
            slides = context.Presentation.Slides;
            return slides.Count;
        }
        finally
        {
            ComUtilities.Release(ref slides);
        }
    }

    private sealed class BeforeExecuteBatch(
        IPresentationBatch inner,
        Action beforeExecute) : IPresentationBatch
    {
        private int _beforeExecutePending = 1;

        public string PresentationPath => inner.PresentationPath;
        public bool HasTimedOutOperation => inner.HasTimedOutOperation;
        public int? PowerPointProcessId => inner.PowerPointProcessId;
        public PowerPointProcessIdentity? PowerPointProcessIdentity => inner.PowerPointProcessIdentity;
        public TimeSpan OperationTimeout => inner.OperationTimeout;

        public void Execute(
            Action<PresentationContext, CancellationToken> operation,
            CancellationToken cancellationToken = default)
        {
            RunBeforeExecute();
            inner.Execute(operation, cancellationToken);
        }

        public T Execute<T>(
            Func<PresentationContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default)
        {
            RunBeforeExecute();
            return inner.Execute(operation, cancellationToken);
        }

        public void Save(CancellationToken cancellationToken = default) =>
            inner.Save(cancellationToken);

        public void UpdatePresentationPath(string presentationPath) =>
            inner.UpdatePresentationPath(presentationPath);

        public bool IsPowerPointProcessAlive() =>
            inner.IsPowerPointProcessAlive();

        public void Dispose()
        {
        }

        private void RunBeforeExecute()
        {
            if (Interlocked.Exchange(ref _beforeExecutePending, 0) == 1)
            {
                beforeExecute();
            }
        }
    }

    [Fact]
    public void Create_SavesRealPptxFile_ThatPowerPointCanReopen()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            var result = _commands.Create(path);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.True(File.Exists(path), "Create() must produce a real .pptx file on disk.");
            Assert.True(new FileInfo(path).Length > 0, "The saved .pptx must not be empty.");

            // Round-trip: open the file we just created with a fresh PowerPoint COM session
            // and verify it is a valid, readable presentation with the default single slide.
            using var batch = PresentationSession.BeginBatch(path);
            int slideCount = batch.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
            Assert.Equal(1, slideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void CreateNew_ShowsRealWindow_BecauseChartInPlaceActivationRequiresIt()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            using var batch = PresentationSession.CreateNew(path);
            var appState = batch.Execute((ctx, ct) => ctx.Presentation.Application.WindowState);
            int windowCount = batch.Execute((ctx, ct) => ctx.Presentation.Windows.Count);

            // PowerPoint does not support hiding its application window (Application.Visible =
            // False throws), and two workarounds — minimizing (ppWindowMinimized) and moving the
            // window off-screen while Normal — were both tried and both broke Chart.ChartData's
            // in-place activation of its embedded Excel workbook (get-chart-data read back 0
            // categories/series in both cases). So the window is always a real, visible,
            // ppWindowNormal document window regardless of the `show` flag — see
            // PresentationBatch.RunStaThread for the full rationale.
            Assert.Equal(1, windowCount);
            Assert.Equal(Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowNormal, appState);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Save_PersistsChanges_VisibleAfterReopen()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);

            using (var batch = PresentationSession.BeginBatch(path))
            {
                batch.Execute((ctx, ct) =>
                {
                    ctx.Presentation.Slides.Add(2, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
                    return 0;
                });

                var saveResult = _commands.Save(batch);
                Assert.True(saveResult.Success);
                Assert.Null(saveResult.ErrorMessage);
            }

            using var reopened = PresentationSession.BeginBatch(path);
            int slideCount = reopened.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
            Assert.Equal(2, slideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_ExistingFile_ReturnsSuccess_WithPresentationPath()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);

            var result = _commands.Open(path);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(Path.GetFullPath(path), result.PresentationPath);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_MissingFile_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        // Deliberately never create the file — Open() must fail gracefully (Rule 1b) without
        // ever starting PowerPoint, not throw.

        var result = _commands.Open(path);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Null(result.PresentationPath);
    }

    [Fact]
    public void Open_ThenEdit_RequiresCallerToBeginItsOwnBatch()
    {
        // Open() proves the file opens and closes it again — it does NOT hand back a live
        // session. Callers who want to edit must call PresentationSession.BeginBatch themselves.
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            var openResult = _commands.Open(path);
            Assert.True(openResult.Success);

            using var batch = PresentationSession.BeginBatch(openResult.PresentationPath!);
            batch.Execute((ctx, ct) =>
            {
                ctx.Presentation.Slides.Add(2, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
                return 0;
            });
            batch.Save();

            using var reopened = PresentationSession.BeginBatch(path);
            int slideCount = reopened.Execute((ctx, ct) => ctx.Presentation.Slides.Count);
            Assert.Equal(2, slideCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Dispose_QuitsPowerPoint_ProcessEventuallyExits()
    {
        // Exercises PresentationShutdownService's resilient close/quit + process-exit polling
        // (invoked internally from PresentationBatch's STA-thread cleanup on Dispose()).
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);

            var batch = PresentationSession.BeginBatch(path);
            int? processId;
            try
            {
                processId = batch.PowerPointProcessId;
                Assert.True(processId.HasValue, "Expected to capture a PowerPoint process ID for the shutdown-polling test.");
                Assert.True(batch.IsPowerPointProcessAlive(), "PowerPoint process should be alive while the batch is open.");
            }
            finally
            {
                batch.Dispose();
            }

            // PresentationShutdownService tolerates PowerPoint's documented ~90-100s post-Quit()
            // Office-cleanup lingering and never force-kills within that window, so poll generously here
            // rather than asserting an immediate exit.
            bool exited = false;
            var deadline = DateTime.UtcNow.AddSeconds(150);
            while (DateTime.UtcNow < deadline)
            {
                if (!batch.IsPowerPointProcessAlive())
                {
                    exited = true;
                    break;
                }
                Thread.Sleep(1000);
            }

            Assert.True(exited, $"PowerPoint process {processId} should eventually exit after Dispose() (within the shutdown service's grace period).");
        }
        finally
        {
            File.Delete(path);
        }
    }

}
