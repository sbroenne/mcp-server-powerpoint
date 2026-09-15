using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Layout;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for shape commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Shape")]
public class ShapeCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _commands = new();
    private readonly TextFrameCommands _textFrameCommands = new();

    public ShapeCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddRectangle_IncreasesShapeCount_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch, 1);
        Assert.Equal(1, countResult.ShapeCount);
    }

    [Fact]
    public void AddTextBox_SetsText_ReadableAfterReopen()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddTextBox(batch, 1, 0f, 0f, 200f, 40f, "Hello PowerPoint");
        Assert.True(result.Success);
        Assert.Equal(1, result.ShapeIndex);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        string text = batch.Execute((ctx, ct) =>
            ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Text);
        Assert.Equal("Hello PowerPoint", text);
    }

    [Fact]
    public void SetPositionAndSize_UpdatesShapeGeometry()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var posResult = _commands.SetPosition(batch, 1, 1, 123f, 456f);
        Assert.True(posResult.Success);
        Assert.Equal(123f, posResult.Left);
        Assert.Equal(456f, posResult.Top);

        var sizeResult = _commands.SetSize(batch, 1, 1, 300f, 200f);
        Assert.True(sizeResult.Success);
        Assert.Equal(300f, sizeResult.Width);
        Assert.Equal(200f, sizeResult.Height);
    }

    [Fact]
    public void Delete_RemovesShape_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        var deleteResult = _commands.Delete(batch, 1, 1);

        Assert.True(deleteResult.Success);
        Assert.Equal(0, deleteResult.ShapeCount);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        var countResult = _commands.GetCount(batch, 1);
        Assert.Equal(0, countResult.ShapeCount);
    }

    [Fact]
    public void AddRectangle_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddRectangle(batch, 99, 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddAutoShape_WithOval_IncreasesShapeCount_AndEchoesShapeTypeName()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeOval", 10f, 20f, 100f, 50f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal("msoShapeOval", result.ShapeTypeName);
    }

    [Fact]
    public void AddAutoShape_WithRightArrow_IncreasesShapeCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeRightArrow", 0f, 0f, 80f, 40f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeCount);
    }

    [Fact]
    public void AddAutoShape_WithUnrecognizedShapeType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 1, "msoShapeDoesNotExist", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddAutoShape_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddAutoShape(batch, 99, "msoShapeOval", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddLine_IncreasesShapeCount_AndEchoesCoordinates()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddLine(batch, 1, 10f, 20f, 200f, 100f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal(10f, result.BeginX);
        Assert.Equal(20f, result.BeginY);
        Assert.Equal(200f, result.EndX);
        Assert.Equal(100f, result.EndY);
    }

    [Fact]
    public void AddLine_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddLine(batch, 99, 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddConnector_WithStraightType_IncreasesShapeCount_AndEchoesTypeAndCoordinates()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorStraight", 5f, 5f, 150f, 75f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeIndex);
        Assert.Equal(1, result.ShapeCount);
        Assert.Equal("msoConnectorStraight", result.ConnectorTypeName);
        Assert.Equal(5f, result.BeginX);
        Assert.Equal(5f, result.BeginY);
        Assert.Equal(150f, result.EndX);
        Assert.Equal(75f, result.EndY);
    }

    [Fact]
    public void AddConnector_WithElbowType_IncreasesShapeCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorElbow", 0f, 0f, 100f, 100f);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ShapeCount);
    }

    [Fact]
    public void AddConnector_WithUnrecognizedConnectorType_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 1, "msoConnectorDoesNotExist", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddConnector_WithInvalidSlideIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddConnector(batch, 99, "msoConnectorStraight", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void AddAttachedConnector_RemainsAttachedWhenShapesMove_AndPersistsAfterReopen()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 40f, 80f, 100f, 60f);
        _commands.AddRectangle(batch, 1, 300f, 220f, 120f, 70f);
        _commands.SetName(batch, 1, 1, "Connector Begin");
        _commands.SetName(batch, 1, 2, "Connector End");

        var result = _commands.AddAttachedConnector(
            batch, 1, "msoConnectorStraight", 1, 2, 2, 4);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(3, result.ShapeIndex);
        Assert.Equal(3, result.ShapeCount);
        Assert.Equal("msoConnectorStraight", result.ConnectorTypeName);

        int connectorIndex = Assert.IsType<int>(result.ShapeIndex);
        var initial = ReadAttachmentState(batch, connectorIndex);
        Assert.True(initial.BeginConnected);
        Assert.True(initial.EndConnected);
        Assert.Equal("Connector Begin", initial.BeginShapeName);
        Assert.Equal("Connector End", initial.EndShapeName);
        Assert.Equal(2, initial.BeginConnectionSite);
        Assert.Equal(4, initial.EndConnectionSite);

        _commands.SetPosition(batch, 1, 1, 120f, 160f);
        _commands.SetPosition(batch, 1, 2, 500f, 300f);

        var moved = ReadAttachmentState(batch, connectorIndex);
        Assert.Equal(initial.BeginShapeName, moved.BeginShapeName);
        Assert.Equal(initial.EndShapeName, moved.EndShapeName);
        Assert.Equal(initial.BeginConnectionSite, moved.BeginConnectionSite);
        Assert.Equal(initial.EndConnectionSite, moved.EndConnectionSite);

        _presentationCommands.Save(batch);
        _fixture.ReopenCurrentPresentation();

        var reopened = ReadAttachmentState(batch, connectorIndex);
        Assert.True(reopened.BeginConnected);
        Assert.True(reopened.EndConnected);
        Assert.Equal(initial.BeginShapeName, reopened.BeginShapeName);
        Assert.Equal(initial.EndShapeName, reopened.EndShapeName);
        Assert.Equal(initial.BeginConnectionSite, reopened.BeginConnectionSite);
        Assert.Equal(initial.EndConnectionSite, reopened.EndConnectionSite);
    }

    [Theory]
    [InlineData("msoConnectorDoesNotExist", 1, 1, 2, 1, "not a recognized MsoConnectorType")]
    [InlineData("msoConnectorStraight", 0, 1, 2, 1, "Shape index 0 is out of range")]
    [InlineData("msoConnectorStraight", 1, 1, 3, 1, "Shape index 3 is out of range")]
    [InlineData("msoConnectorStraight", 1, 0, 2, 1, "Begin connection site 0 is out of range")]
    [InlineData("msoConnectorStraight", 1, 1, 2, 5, "End connection site 5 is out of range")]
    public void AddAttachedConnector_WithInvalidEndpoint_ReturnsFailureWithoutAddingShape(
        string connectorType,
        int beginShapeIndex,
        int beginConnectionSite,
        int endShapeIndex,
        int endConnectionSite,
        string expectedError)
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 40f, 80f, 100f, 60f);
        _commands.AddRectangle(batch, 1, 300f, 220f, 120f, 70f);

        var result = _commands.AddAttachedConnector(
            batch,
            1,
            connectorType,
            beginShapeIndex,
            beginConnectionSite,
            endShapeIndex,
            endConnectionSite);

        Assert.False(result.Success);
        Assert.Contains(expectedError, result.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(2, _commands.GetCount(batch, 1).ShapeCount);
    }

    private static (bool BeginConnected, bool EndConnected, string BeginShapeName, string EndShapeName, int BeginConnectionSite, int EndConnectionSite)
        ReadAttachmentState(ComInterop.Session.IPresentationBatch batch, int connectorIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? connector = null;
            PowerPoint.ConnectorFormat? connectorFormat = null;
            PowerPoint.Shape? beginShape = null;
            PowerPoint.Shape? endShape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                slide = slides[1];
                shapes = slide.Shapes;
                connector = shapes[connectorIndex];
                connectorFormat = connector.ConnectorFormat;
                beginShape = connectorFormat.BeginConnectedShape;
                endShape = connectorFormat.EndConnectedShape;
                return (
                    Convert.ToInt32(connectorFormat.BeginConnected, CultureInfo.InvariantCulture) != 0,
                    Convert.ToInt32(connectorFormat.EndConnected, CultureInfo.InvariantCulture) != 0,
                    beginShape.Name,
                    endShape.Name,
                    connectorFormat.BeginConnectionSite,
                    connectorFormat.EndConnectionSite);
            }
            finally
            {
                if (endShape is not null) ComInterop.ComUtilities.Release(ref endShape);
                if (beginShape is not null) ComInterop.ComUtilities.Release(ref beginShape);
                if (connectorFormat is not null) ComInterop.ComUtilities.Release(ref connectorFormat);
                if (connector is not null) ComInterop.ComUtilities.Release(ref connector);
                if (shapes is not null) ComInterop.ComUtilities.Release(ref shapes);
                if (slide is not null) ComInterop.ComUtilities.Release(ref slide);
                if (slides is not null) ComInterop.ComUtilities.Release(ref slides);
            }
        });
    }

    [Fact]
    public void SetFill_AndGetFill_RoundTripsColor()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetFill(batch, 1, 1, 255, 0, 0);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);

        var getResult = _commands.GetFill(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(255, getResult.ColorRgb);
    }

    [Fact]
    public void SetFill_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetFill(batch, 1, 99, 255, 0, 0);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetLine_AndGetLine_RoundTripsColorWeightDashStyleAndVisibility()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetLine(batch, 1, 1, red: 0, green: 255, blue: 0, weight: 3f, dashStyle: "msoLineDash", visible: true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(65280, setResult.ColorRgb);
        Assert.Equal(3f, setResult.LineWeight);
        Assert.Equal("msoLineDash", setResult.DashStyleName);
        Assert.True(setResult.Visible);

        var getResult = _commands.GetLine(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(65280, getResult.ColorRgb);
        Assert.Equal(3f, getResult.LineWeight);
        Assert.Equal("msoLineDash", getResult.DashStyleName);
        Assert.True(getResult.Visible);
    }

    [Fact]
    public void SetLine_WithUnrecognizedDashStyle_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.SetLine(batch, 1, 1, dashStyle: "msoLineDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetLine_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetLine(batch, 1, 99, weight: 3f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void CopyFormatting_TransfersAppearanceWithoutReplacingContentOrGeometry_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 10f, 20f, 120f, 40f);
        _commands.AddRectangle(batch, 1, 200f, 220f, 180f, 70f);
        Assert.True(_textFrameCommands.SetText(batch, 1, 1, "Source content").Success);
        Assert.True(_textFrameCommands.SetText(batch, 1, 2, "Target content").Success);
        _commands.SetFill(batch, 1, 1, 12, 34, 56);
        _commands.SetLine(batch, 1, 1, red: 78, green: 90, blue: 123, weight: 4f, dashStyle: "msoLineDash", visible: true);
        _commands.SetFill(batch, 1, 2, 200, 210, 220);

        var result = _commands.CopyFormatting(batch, 1, 1, 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.ShapeIndex);
        Assert.Equal(0x38220C, _commands.GetFill(batch, 1, 2).ColorRgb);
        var targetLine = _commands.GetLine(batch, 1, 2);
        Assert.Equal(0x7B5A4E, targetLine.ColorRgb);
        Assert.Equal(4f, targetLine.LineWeight);
        Assert.Equal("msoLineDash", targetLine.DashStyleName);

        var targetState = batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? target = null;
            try
            {
                slides = ctx.Presentation.Slides;
                slide = slides[1];
                shapes = slide.Shapes;
                target = shapes[2];
                return (target.Left, target.Top, target.Width, target.Height);
            }
            finally
            {
                if (target is not null) ComInterop.ComUtilities.Release(ref target);
                if (shapes is not null) ComInterop.ComUtilities.Release(ref shapes);
                if (slide is not null) ComInterop.ComUtilities.Release(ref slide);
                if (slides is not null) ComInterop.ComUtilities.Release(ref slides);
            }
        });
        Assert.Equal("Target content", _textFrameCommands.GetText(batch, 1, 2).Text);
        Assert.Equal(200f, targetState.Left);
        Assert.Equal(220f, targetState.Top);
        Assert.Equal(180f, targetState.Width);
        Assert.Equal(70f, targetState.Height);

        _presentationCommands.Save(batch);
        _fixture.ReopenCurrentPresentation();

        Assert.Equal(0x38220C, _commands.GetFill(batch, 1, 2).ColorRgb);
        Assert.Equal("Target content", _textFrameCommands.GetText(batch, 1, 2).Text);
    }

    [Fact]
    public async Task CopyFormatting_WaitsForGlobalFormattingClipboardLock()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 10f, 20f, 120f, 40f);
        _commands.AddRectangle(batch, 1, 200f, 220f, 180f, 70f);
        _commands.SetFill(batch, 1, 1, 12, 34, 56);
        _commands.SetFill(batch, 1, 2, 200, 210, 220);

        using var lockAcquired = new ManualResetEventSlim();
        using var releaseLock = new ManualResetEventSlim();
        Exception? holderFailure = null;
        var lockHolder = new Thread(() =>
            FormattingClipboardTestLock.Hold(lockAcquired, releaseLock, ref holderFailure))
        {
            IsBackground = true
        };
        lockHolder.Start();
        Task<ShapeOperationResult>? copyTask = null;
        try
        {
            Assert.True(lockAcquired.Wait(TimeSpan.FromSeconds(15)));
            Assert.Null(holderFailure);

            copyTask = Task.Run(() => _commands.CopyFormatting(batch, 1, 1, 2));
            var completedTask = await Task.WhenAny(copyTask, Task.Delay(TimeSpan.FromSeconds(3)));
            Assert.False(
                ReferenceEquals(copyTask, completedTask),
                "CopyFormatting completed while another owner held the global formatting clipboard lock.");
        }
        finally
        {
            releaseLock.Set();
            await Task.Run(lockHolder.Join);
        }

        // Bounded like the wait above: a stalled transfer must fail this test, not hang the host.
        Assert.Same(copyTask, await Task.WhenAny(copyTask, Task.Delay(TimeSpan.FromMinutes(2))));
        var result = await copyTask;
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(0x38220C, _commands.GetFill(batch, 1, 2).ColorRgb);
    }

    [Fact]
    public void CopyFormatting_WhenClipboardLockStaysHeld_TimesOutAndLeavesSessionUsable()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 10f, 20f, 120f, 40f);
        _commands.AddRectangle(batch, 1, 200f, 220f, 180f, 70f);
        _commands.SetFill(batch, 1, 1, 12, 34, 56);
        _commands.SetFill(batch, 1, 2, 200, 210, 220);

        using var lockAcquired = new ManualResetEventSlim();
        using var releaseLock = new ManualResetEventSlim();
        Exception? holderFailure = null;
        var lockHolder = new Thread(() =>
            FormattingClipboardTestLock.Hold(lockAcquired, releaseLock, ref holderFailure))
        {
            IsBackground = true
        };
        lockHolder.Start();
        try
        {
            Assert.True(lockAcquired.Wait(TimeSpan.FromSeconds(15)));
            Assert.Null(holderFailure);

            var timeout = Assert.Throws<TimeoutException>(() => _commands.CopyFormatting(batch, 1, 1, 2));
            Assert.Contains("formatting clipboard", timeout.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            releaseLock.Set();
            lockHolder.Join();
        }

        // The wait has to expire before batch.Execute's own operation timeout, otherwise the STA
        // thread stays blocked on the lock and the session is poisoned by mere contention.
        Assert.False(batch.HasTimedOutOperation);
        Assert.Equal(0xDCD2C8, _commands.GetFill(batch, 1, 2).ColorRgb);
    }

    [Fact]
    public void CopyFormatting_OnPoisonedSession_FailsFastInsteadOfWaitingForTheLock()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 10f, 20f, 120f, 40f);
        _commands.AddRectangle(batch, 1, 200f, 220f, 180f, 70f);

        using var lockAcquired = new ManualResetEventSlim();
        using var releaseLock = new ManualResetEventSlim();
        Exception? holderFailure = null;
        var lockHolder = new Thread(() =>
            FormattingClipboardTestLock.Hold(lockAcquired, releaseLock, ref holderFailure))
        {
            IsBackground = true
        };
        lockHolder.Start();
        try
        {
            Assert.True(lockAcquired.Wait(TimeSpan.FromSeconds(15)));
            Assert.Null(holderFailure);

            var poisoned = new PoisonedSessionBatch(batch);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            Assert.Throws<TimeoutException>(() => _commands.CopyFormatting(poisoned, 1, 1, 2));
            stopwatch.Stop();

            // Without the preflight this would first wait out the whole formatting-lock timeout
            // (half the batch's operation timeout) before reporting the dead session.
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(10),
                $"A poisoned session took {stopwatch.Elapsed} to fail, so it waited on the formatting lock first.");
        }
        finally
        {
            releaseLock.Set();
            lockHolder.Join();
        }
    }

    /// <summary>
    /// Wraps a real batch but reports the poisoned-session state that <c>PresentationBatch</c>
    /// enters after an operation times out, so the fail-fast path can be exercised without
    /// actually breaking the shared PowerPoint session.
    /// </summary>
    private sealed class PoisonedSessionBatch(IPresentationBatch inner) : IPresentationBatch
    {
        public string PresentationPath => inner.PresentationPath;
        public bool HasTimedOutOperation => true;
        public int? PowerPointProcessId => inner.PowerPointProcessId;
        public PowerPointProcessIdentity? PowerPointProcessIdentity => inner.PowerPointProcessIdentity;
        public TimeSpan OperationTimeout => inner.OperationTimeout;

        public void Execute(
            Action<PresentationContext, CancellationToken> operation,
            CancellationToken cancellationToken = default) =>
            throw new TimeoutException("A previous operation timed out for this presentation.");

        public T Execute<T>(
            Func<PresentationContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default) =>
            throw new TimeoutException("A previous operation timed out for this presentation.");

        public void Save(CancellationToken cancellationToken = default) => inner.Save(cancellationToken);

        public void UpdatePresentationPath(string presentationPath) =>
            inner.UpdatePresentationPath(presentationPath);

        public bool IsPowerPointProcessAlive() => false;

        public void Dispose()
        {
        }
    }

    [Theory]
    [InlineData(0, 2)]
    [InlineData(99, 2)]
    [InlineData(1, 0)]
    [InlineData(1, 99)]
    public void CopyFormatting_WithInvalidShapeIndex_ReturnsFailureWithoutMutatingTarget(
        int sourceShapeIndex,
        int targetShapeIndex)
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 10f, 20f, 120f, 40f);
        _commands.AddRectangle(batch, 1, 200f, 220f, 180f, 70f);
        _commands.SetFill(batch, 1, 1, 12, 34, 56);
        _commands.SetFill(batch, 1, 2, 200, 210, 220);

        var result = _commands.CopyFormatting(batch, 1, sourceShapeIndex, targetShapeIndex);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(0xDCD2C8, _commands.GetFill(batch, 1, 2).ColorRgb);
    }

    [Fact]
    public void SetRotation_AndGetRotation_RoundTripsDegrees()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetRotation(batch, 1, 1, 45f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(45f, setResult.Rotation);

        var getResult = _commands.GetRotation(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(45f, getResult.Rotation);
    }

    [Fact]
    public void SetRotation_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetRotation(batch, 1, 99, 45f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void Flip_WithHorizontalDirection_Succeeds()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.Flip(batch, 1, 1, "horizontal");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("horizontal", result.FlipDirection);
    }

    [Fact]
    public void Flip_WithVerticalDirection_Succeeds()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.Flip(batch, 1, 1, "vertical");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("vertical", result.FlipDirection);
    }

    [Fact]
    public void Flip_WithUnrecognizedDirection_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.Flip(batch, 1, 1, "diagonal");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetZOrder_WithBringToFront_Succeeds()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        _commands.AddRectangle(batch, 1, 10f, 10f, 50f, 50f);

        var result = _commands.SetZOrder(batch, 1, 1, "bring-to-front");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("bring-to-front", result.ZOrderCommand);
    }

    [Fact]
    public void SetZOrder_WithUnrecognizedCommand_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.SetZOrder(batch, 1, 1, "bring-in-front-of-text");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetShadow_AndGetShadow_RoundTripsVisibility()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetShadow(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Visible);

        var getResult = _commands.GetShadow(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Visible);

        var offResult = _commands.SetShadow(batch, 1, 1, false);
        Assert.True(offResult.Success, offResult.ErrorMessage);
        Assert.False(offResult.Visible);
    }

    [Fact]
    public void SetShadow_WithParameters_RoundTripsColorAndFormatting()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetShadow(batch, 1, 1, true, red: 255, green: 0, blue: 0, transparency: 0.25f, blur: 4f, offsetX: 4f, offsetY: 5f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);
        Assert.Equal(0.25f, setResult.Transparency);
        // PowerPoint round-trips points through EMUs internally, so exact float equality is not
        // guaranteed for shadow Blur/OffsetX/OffsetY — assert within a small tolerance instead.
        Assert.InRange(setResult.Blur!.Value, 3f, 7f);
        Assert.InRange(setResult.OffsetX!.Value, 3.9f, 4.1f);
        Assert.InRange(setResult.OffsetY!.Value, 4.9f, 5.1f);

        var getResult = _commands.GetShadow(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Visible);
        Assert.Equal(255, getResult.ColorRgb);
        Assert.Equal(0.25f, getResult.Transparency);
        Assert.InRange(getResult.Blur!.Value, 3f, 7f);
        Assert.InRange(getResult.OffsetX!.Value, 3.9f, 4.1f);
        Assert.InRange(getResult.OffsetY!.Value, 4.9f, 5.1f);
    }

    [Fact]
    public void SetGlow_AndGetGlow_RoundTripsColorRadiusAndTransparency()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetGlow(batch, 1, 1, red: 0, green: 255, blue: 0, radius: 8f, transparency: 0.3f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(0x00FF00, setResult.ColorRgb);
        Assert.InRange(setResult.GlowRadius!.Value, 7.9f, 8.1f);
        Assert.Equal(0.3f, setResult.Transparency);

        var getResult = _commands.GetGlow(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(0x00FF00, getResult.ColorRgb);
        Assert.InRange(getResult.GlowRadius!.Value, 7.9f, 8.1f);
        Assert.Equal(0.3f, getResult.Transparency);
    }

    [Fact]
    public void SetReflection_AndGetReflection_RoundTripsVisibilityAndFormatting()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetReflection(batch, 1, 1, true, transparency: 0.6f, size: 40f, blur: 2f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Visible);
        Assert.Equal(0.6f, setResult.Transparency);
        Assert.InRange(setResult.ReflectionSize!.Value, 39.9f, 40.1f);
        Assert.InRange(setResult.Blur!.Value, 1.9f, 2.1f);

        var getResult = _commands.GetReflection(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Visible);
        Assert.Equal(0.6f, getResult.Transparency);
        Assert.InRange(getResult.ReflectionSize!.Value, 39.9f, 40.1f);
        Assert.InRange(getResult.Blur!.Value, 1.9f, 2.1f);

        var offResult = _commands.SetReflection(batch, 1, 1, false);
        Assert.True(offResult.Success, offResult.ErrorMessage);
        Assert.False(offResult.Visible);

        var getOffResult = _commands.GetReflection(batch, 1, 1);
        Assert.True(getOffResult.Success, getOffResult.ErrorMessage);
        Assert.False(getOffResult.Visible);
    }

    [Fact]
    public void SetSoftEdge_AndGetSoftEdge_RoundTripsRadius()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetSoftEdge(batch, 1, 1, 5f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.InRange(setResult.SoftEdgeRadius!.Value, 4.9f, 5.1f);

        var getResult = _commands.GetSoftEdge(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.InRange(getResult.SoftEdgeRadius!.Value, 4.9f, 5.1f);
    }

    [Fact]
    public void SetBevel_AndGetBevel_RoundTripsTypeDepthAndInset()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetBevel(batch, 1, 1, "msoBevelCircle", depth: 7f, inset: 8f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("msoBevelCircle", setResult.BevelTypeName);
        Assert.InRange(setResult.BevelDepth!.Value, 6.9f, 7.1f);
        Assert.InRange(setResult.BevelInset!.Value, 7.9f, 8.1f);

        var getResult = _commands.GetBevel(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("msoBevelCircle", getResult.BevelTypeName);
        Assert.InRange(getResult.BevelDepth!.Value, 6.9f, 7.1f);
        Assert.InRange(getResult.BevelInset!.Value, 7.9f, 8.1f);
    }

    [Fact]
    public void SetBevel_WithUnrecognizedTypeName_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.SetBevel(batch, 1, 1, "msoBevelNotARealType");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void Group_TwoShapes_ReducesShapeCountByOne_AndUngroupRestoresIt()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        _commands.AddRectangle(batch, 1, 60f, 0f, 50f, 50f);

        var groupResult = _commands.Group(batch, 1, [1, 2]);
        Assert.True(groupResult.Success, groupResult.ErrorMessage);
        Assert.Equal(1, groupResult.ShapeCount);

        var ungroupResult = _commands.Ungroup(batch, 1, 1);
        Assert.True(ungroupResult.Success, ungroupResult.ErrorMessage);
        Assert.Equal(2, ungroupResult.UngroupedShapeCount);
        Assert.Equal(2, ungroupResult.ShapeCount);
    }

    [Fact]
    public void Group_WithFewerThanTwoIndexes_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.Group(batch, 1, [1]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void Group_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.Group(batch, 1, [1, 99]);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetName_AndGetName_RoundTripsName()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetName(batch, 1, 1, "MyRectangle");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("MyRectangle", setResult.Name);

        var getResult = _commands.GetName(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("MyRectangle", getResult.Name);
    }

    [Fact]
    public void SetName_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetName(batch, 1, 99, "MyRectangle");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAltText_AndGetAltText_RoundTripsText()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetAltText(batch, 1, 1, "A red rectangle");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("A red rectangle", setResult.AltText);

        var getResult = _commands.GetAltText(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("A red rectangle", getResult.AltText);
    }

    [Fact]
    public void SetAltText_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetAltText(batch, 1, 99, "A red rectangle");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetHyperlink_AndGetHyperlink_RoundTripsAddress()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetHyperlink(batch, 1, 1, "https://example.com");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.HasHyperlink);
        Assert.Equal("https://example.com/", setResult.HyperlinkAddress);

        var getResult = _commands.GetHyperlink(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.HasHyperlink);
        Assert.Equal("https://example.com/", getResult.HyperlinkAddress);
    }

    [Fact]
    public void LinkedPicture_LinkLifecycle_PersistsThenBreaksAndRetainsImage()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            var imageCommands = new ImageCommands();
            var addResult = imageCommands.AddPicture(
                batch, 1, imagePath, 10f, 10f, 100f, 100f,
                linkToFile: true, saveWithDocument: false);
            Assert.True(addResult.Success, addResult.ErrorMessage);

            var info = _commands.GetLinkInfo(batch, 1, 1);
            Assert.True(info.Success, info.ErrorMessage);
            Assert.True(info.HasLink);
            Assert.Equal(Path.GetFullPath(imagePath), info.LinkSourceFullName);
            Assert.Null(info.LinkAutoUpdate);

            var autoUpdateException = Record.Exception(() =>
            {
                var autoUpdate = _commands.SetLinkAutoUpdate(batch, 1, 1, false);
                Assert.True(autoUpdate.Success, autoUpdate.ErrorMessage);
                Assert.False(autoUpdate.LinkAutoUpdate);
            });
            AssertSupportedAutoUpdateOrKnownPowerPointError(autoUpdateException);

            var update = _commands.UpdateLink(batch, 1, 1);
            Assert.True(update.Success, update.ErrorMessage);

            _presentationCommands.Save(batch);
            _fixture.ReopenCurrentPresentation();

            var reopenedInfo = _commands.GetLinkInfo(batch, 1, 1);
            Assert.True(reopenedInfo.Success, reopenedInfo.ErrorMessage);
            Assert.True(reopenedInfo.HasLink);
            Assert.Equal(Path.GetFullPath(imagePath), reopenedInfo.LinkSourceFullName);
            Assert.Null(reopenedInfo.LinkAutoUpdate);

            var breakResult = _commands.BreakLink(batch, 1, 1);
            Assert.True(breakResult.Success, breakResult.ErrorMessage);
            Assert.False(breakResult.HasLink);

            File.Delete(imagePath);
            _presentationCommands.Save(batch);
            _fixture.ReopenCurrentPresentation();

            var countResult = _commands.GetCount(batch, 1);
            Assert.True(countResult.Success, countResult.ErrorMessage);
            Assert.Equal(1, countResult.ShapeCount);

            var afterBreak = _commands.GetLinkInfo(batch, 1, 1);
            Assert.False(afterBreak.Success);
            Assert.Contains("not linked", afterBreak.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    private static void AssertSupportedAutoUpdateOrKnownPowerPointError(Exception? exception)
    {
        if (exception is null)
        {
            return;
        }

        var comException = Assert.IsType<System.Runtime.InteropServices.COMException>(exception);
        Assert.Equal(unchecked((int)0x80048240), comException.HResult);
    }

    [Fact]
    public void LinkOperations_OnOrdinaryShape_ReturnExpectedValidationFailures()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        var results = new[]
        {
            _commands.GetLinkInfo(batch, 1, 1),
            _commands.UpdateLink(batch, 1, 1),
            _commands.BreakLink(batch, 1, 1),
            _commands.SetLinkAutoUpdate(batch, 1, 1, true)
        };

        foreach (var result in results)
        {
            Assert.False(result.Success);
            Assert.Contains("not linked", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void GetLinkInfo_WithZeroShapeIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var result = _commands.GetLinkInfo(_fixture.Batch, 1, 0);

        Assert.False(result.Success);
        Assert.Contains("Shape index 0", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void SetHyperlink_WithScreenTip_RoundTripsScreenTip()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var setResult = _commands.SetHyperlink(batch, 1, 1, "https://example.com", screenTip: "Visit Example");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Visit Example", setResult.HyperlinkScreenTip);

        var getResult = _commands.GetHyperlink(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("Visit Example", getResult.HyperlinkScreenTip);
    }

    [Fact]
    public void GetHyperlink_OnShapeWithoutHyperlink_ReturnsHasHyperlinkFalse()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);

        var result = _commands.GetHyperlink(batch, 1, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.HasHyperlink);
        Assert.Null(result.HyperlinkAddress);
    }

    [Fact]
    public void RemoveHyperlink_ClearsPreviouslySetHyperlink()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _commands.AddRectangle(batch, 1, 0f, 0f, 50f, 50f);
        _commands.SetHyperlink(batch, 1, 1, "https://example.com");

        var removeResult = _commands.RemoveHyperlink(batch, 1, 1);
        Assert.True(removeResult.Success, removeResult.ErrorMessage);
        Assert.False(removeResult.HasHyperlink);

        var getResult = _commands.GetHyperlink(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.HasHyperlink);
        Assert.Null(getResult.HyperlinkAddress);
    }

    [Fact]
    public void SetHyperlink_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetHyperlink(batch, 1, 99, "https://example.com");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ListPlaceholders_AndSetPlaceholderText_RoundTripTitlePlaceholder()
    {
        _fixture.CreateFreshPresentation();
        var layoutResult = new LayoutCommands().SetLayout(_fixture.Batch, 1, "ppLayoutTitleOnly");
        Assert.True(layoutResult.Success, layoutResult.ErrorMessage);

        var listed = _commands.ListPlaceholders(_fixture.Batch, 1);

        Assert.True(listed.Success, listed.ErrorMessage);
        var title = Assert.Single(listed.Placeholders!);
        Assert.True(title.ShapeIndex >= 1);
        Assert.Equal("ppPlaceholderTitle", title.PlaceholderType);

        var setResult = _commands.SetPlaceholderText(
            _fixture.Batch, 1, title.ShapeIndex, "Release title");

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Null(setResult.ErrorMessage);

        var saveResult = _presentationCommands.Save(_fixture.Batch);
        Assert.True(saveResult.Success, saveResult.ErrorMessage);
        _fixture.ReopenCurrentPresentation();

        string nativeText = _fixture.Batch.Execute((ctx, ct) =>
            ctx.Presentation.Slides[1].Shapes[title.ShapeIndex].TextFrame.TextRange.Text);
        Assert.Equal("Release title", nativeText);
    }

    [Fact]
    public void SetPlaceholderText_OnNonPlaceholder_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var added = _commands.AddTextBox(_fixture.Batch, 1, 10f, 10f, 200f, 40f, "Not a placeholder");
        Assert.True(added.Success, added.ErrorMessage);

        var result = _commands.SetPlaceholderText(
            _fixture.Batch, 1, added.ShapeIndex!.Value, "Replacement text");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
    }

    [Fact]
    public void SetPlaceholderImage_ReplacesNativePicturePlaceholder()
    {
        _fixture.CreateFreshPresentation();
        var layoutResult = new LayoutCommands().SetLayout(
            _fixture.Batch, 1, "ppLayoutPictureWithCaption");
        Assert.True(layoutResult.Success, layoutResult.ErrorMessage);

        var listed = _commands.ListPlaceholders(_fixture.Batch, 1);
        Assert.True(listed.Success, listed.ErrorMessage);
        var picturePlaceholder = Assert.Single(
            listed.Placeholders!,
            placeholder => placeholder.PlaceholderType == "ppPlaceholderPicture");

        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            var result = _commands.SetPlaceholderImage(
                _fixture.Batch, 1, picturePlaceholder.ShapeIndex, imagePath);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);

            var saveResult = _presentationCommands.Save(_fixture.Batch);
            Assert.True(saveResult.Success, saveResult.ErrorMessage);
            _fixture.ReopenCurrentPresentation();

            var after = _commands.ListPlaceholders(_fixture.Batch, 1);
            Assert.True(after.Success, after.ErrorMessage);
            var replaced = Assert.Single(
                after.Placeholders!,
                placeholder => placeholder.PlaceholderType == "ppPlaceholderPicture");
            Assert.True(replaced.HasImage);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void Tags_CrudIsCaseInsensitive_EnumeratesOneBased_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_commands.AddRectangle(_fixture.Batch, 1, 10f, 10f, 100f, 50f).Success);

        var firstSet = _commands.SetTag(_fixture.Batch, 1, 1, " ReviewState ", "MiXeD Value");
        Assert.True(firstSet.Success, firstSet.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", firstSet.TagName);
        Assert.Equal("MiXeD Value", firstSet.TagValue);
        Assert.Equal(1, firstSet.TagCount);

        var updated = _commands.SetTag(_fixture.Batch, 1, 1, " reviewstate ", "Updated Value");
        Assert.True(updated.Success, updated.ErrorMessage);
        Assert.Equal(1, updated.TagCount);

        Assert.True(_commands.SetTag(_fixture.Batch, 1, 1, "Owner", "Alice").Success);
        Assert.True(_commands.SetTag(_fixture.Batch, 1, 1, "ReviewState", "Unspaced Value").Success);

        var get = _commands.GetTag(_fixture.Batch, 1, 1, " ReViEwStAtE ");
        Assert.True(get.Success, get.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", get.TagName);
        Assert.Equal("Updated Value", get.TagValue);
        Assert.Equal(1, get.TagIndex);

        var unspacedGet = _commands.GetTag(_fixture.Batch, 1, 1, "reviewstate");
        Assert.True(unspacedGet.Success, unspacedGet.ErrorMessage);
        Assert.Equal("Unspaced Value", unspacedGet.TagValue);
        Assert.Equal(3, unspacedGet.TagIndex);

        var listed = _commands.ListTags(_fixture.Batch, 1, 1);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(3, listed.TagCount);
        Assert.Equal([1, 2, 3], listed.Tags!.Select(tag => tag.TagIndex));
        Assert.Equal([" REVIEWSTATE ", "OWNER", "REVIEWSTATE"], listed.Tags!.Select(tag => tag.Name));

        Assert.True(_presentationCommands.Save(_fixture.Batch).Success);
        _fixture.ReopenCurrentPresentation();
        Assert.Equal("Updated Value", _commands.GetTag(_fixture.Batch, 1, 1, " reviewstate ").TagValue);

        var deleted = _commands.DeleteTag(_fixture.Batch, 1, 1, " REVIEWSTATE ");
        Assert.True(deleted.Success, deleted.ErrorMessage);
        Assert.Equal(2, deleted.TagCount);
        Assert.False(_commands.GetTag(_fixture.Batch, 1, 1, " reviewstate ").Success);
        Assert.False(_commands.DeleteTag(_fixture.Batch, 1, 1, " reviewstate ").Success);
        Assert.True(_commands.GetTag(_fixture.Batch, 1, 1, "reviewstate").Success);

        Assert.True(_commands.ListTags(_fixture.Batch, 1, 1).Success);
        Assert.Equal(1, _commands.GetCount(_fixture.Batch, 1).ShapeCount);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(99, 1)]
    [InlineData(1, 0)]
    [InlineData(1, 99)]
    public void Tags_WithInvalidOwnerIndex_ReturnFailure(int slideIndex, int shapeIndex)
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_commands.AddRectangle(_fixture.Batch, 1, 10f, 10f, 100f, 50f).Success);

        Assert.False(_commands.SetTag(_fixture.Batch, slideIndex, shapeIndex, "name", "value").Success);
        Assert.False(_commands.GetTag(_fixture.Batch, slideIndex, shapeIndex, "name").Success);
        Assert.False(_commands.ListTags(_fixture.Batch, slideIndex, shapeIndex).Success);
        Assert.False(_commands.DeleteTag(_fixture.Batch, slideIndex, shapeIndex, "name").Success);
        Assert.Equal(1, _commands.GetCount(_fixture.Batch, 1).ShapeCount);
    }
}
