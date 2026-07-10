using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;

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
            // Office-cleanup lingering (Ripley's MCP round-trip finding, .squad/decisions.md
            // 2026-07-01) and never force-kills within that window — so poll generously here
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

    [Theory]
    [InlineData("Title")]
    [InlineData("subject")]
    [InlineData("AUTHOR")]
    [InlineData("Keywords")]
    [InlineData("Comments")]
    [InlineData("Category")]
    [InlineData("Manager")]
    [InlineData("Company")]
    public void SetDocumentProperty_ThenGetDocumentProperty_RoundTrips_CaseInsensitively(string propertyName)
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var setResult = _commands.SetDocumentProperty(batch, propertyName, "Test Value");
            Assert.True(setResult.Success);
            Assert.Null(setResult.ErrorMessage);
            Assert.Equal("Test Value", setResult.PropertyValue);

            var getResult = _commands.GetDocumentProperty(batch, propertyName);
            Assert.True(getResult.Success);
            Assert.Null(getResult.ErrorMessage);
            Assert.Equal("Test Value", getResult.PropertyValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetDocumentProperty_UnsupportedName_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var result = _commands.SetDocumentProperty(batch, "NotARealProperty", "value");

            Assert.False(result.Success);
            Assert.Contains("NotARealProperty", result.ErrorMessage);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetDocumentProperty_UnsupportedName_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var result = _commands.GetDocumentProperty(batch, "NotARealProperty");

            Assert.False(result.Success);
            Assert.Contains("NotARealProperty", result.ErrorMessage);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCustomProperty_ThenGetCustomProperty_RoundTrips()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var setResult = _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
            Assert.True(setResult.Success);
            Assert.Null(setResult.ErrorMessage);

            var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
            Assert.True(getResult.Success);
            Assert.Equal("ABC-123", getResult.PropertyValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCustomProperty_CalledTwice_UpdatesExistingValue_DoesNotDuplicate()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
            var secondSet = _commands.SetCustomProperty(batch, "ProjectCode", "XYZ-999");
            Assert.True(secondSet.Success);

            var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
            Assert.Equal("XYZ-999", getResult.PropertyValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetCustomProperty_NotFound_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var result = _commands.GetCustomProperty(batch, "DoesNotExist");

            Assert.False(result.Success);
            Assert.Contains("DoesNotExist", result.ErrorMessage);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveCustomProperty_AfterSet_RemovesIt_SubsequentGetFails()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
            var removeResult = _commands.RemoveCustomProperty(batch, "ProjectCode");
            Assert.True(removeResult.Success);
            Assert.Null(removeResult.ErrorMessage);

            var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
            Assert.False(getResult.Success);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveCustomProperty_NotFound_ReturnsFailure_NotException()
    {
        string path = CoreTestHelper.CreateUniqueTestFilePath();
        try
        {
            _commands.Create(path);
            using var batch = PresentationSession.BeginBatch(path);

            var result = _commands.RemoveCustomProperty(batch, "DoesNotExist");

            Assert.False(result.Success);
            Assert.Contains("DoesNotExist", result.ErrorMessage);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
