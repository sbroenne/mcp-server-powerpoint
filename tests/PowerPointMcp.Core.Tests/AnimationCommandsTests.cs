using Sbroenne.PowerPointMcp.Core.Animation;
using Sbroenne.PowerPointMcp.Core.Shape;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for animation commands (shape entrance/emphasis/exit effects, slide
/// transitions) against live PowerPoint COM. No mocking. Shares one PowerPoint.Application
/// instance across all [Fact]s in this class via <see cref="SharedPresentationFixture"/> — each
/// test still gets its own freshly-created presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Animation")]
public class AnimationCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly ShapeCommands _shapeCommands = new();
    private readonly AnimationCommands _commands = new();

    public AnimationCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddEffect_OnShape_IncreasesEffectCount_WithExpectedDefaults()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        var result = _commands.AddEffect(batch, 1, 1, "msoAnimEffectFade");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.EffectIndex);
        Assert.Equal(1, result.EffectCount);
        Assert.Equal("msoAnimEffectFade", result.EffectName);
        Assert.False(result.IsExit);
        Assert.Equal("on-click", result.Trigger);
    }

    [Fact]
    public void AddEffect_AsExitWithAfterPreviousTrigger_RoundTripsFlags()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        var result = _commands.AddEffect(batch, 1, 1, "msoAnimEffectFly", isExit: true, trigger: "after-previous");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.IsExit);
        Assert.Equal("after-previous", result.Trigger);
    }

    [Fact]
    public void AddEffect_WithInvalidEffectName_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        var result = _commands.AddEffect(batch, 1, 1, "msoAnimEffectDoesNotExist");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void AddEffect_WithInvalidShapeIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddEffect(batch, 1, 99, "msoAnimEffectFade");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void AddEffect_WithInvalidSlideIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        var result = _commands.AddEffect(batch, 99, 1, "msoAnimEffectFade");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void GetEffectCount_AfterAddingTwoEffects_ReturnsTwo()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);
        _shapeCommands.AddTextBox(batch, 1, 0f, 0f, 200f, 40f, "Hello");

        _commands.AddEffect(batch, 1, 1, "msoAnimEffectFade");
        _commands.AddEffect(batch, 1, 2, "msoAnimEffectFly");

        var result = _commands.GetEffectCount(batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.EffectCount);
    }

    [Fact]
    public void DeleteEffect_RemovesEffect_AndDecrementsCount()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);
        _commands.AddEffect(batch, 1, 1, "msoAnimEffectFade");

        var result = _commands.DeleteEffect(batch, 1, 1);

        Assert.True(result.Success, result.ErrorMessage);

        var countResult = _commands.GetEffectCount(batch, 1);
        Assert.Equal(0, countResult.EffectCount);
    }

    [Fact]
    public void DeleteEffect_WithInvalidIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 10f, 20f, 100f, 50f);

        var result = _commands.DeleteEffect(batch, 1, 1);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void SetTransition_ChangesEntryEffectAndDuration_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetTransition(batch, 1, "ppEffectFade", durationSeconds: 1.5f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppEffectFade", setResult.TransitionName);
        Assert.Equal(1.5f, setResult.DurationSeconds);

        var getResult = _commands.GetTransition(batch, 1);
        Assert.True(getResult.Success);
        Assert.Equal("ppEffectFade", getResult.TransitionName);
        Assert.Equal(1.5f, getResult.DurationSeconds);
    }

    [Fact]
    public void SetTransition_WithAdvanceOnTime_RoundTripsAdvanceSettings()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetTransition(
            batch, 1, "ppEffectPushLeft", advanceOnClick: false, advanceOnTime: true, advanceTimeSeconds: 5f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.False(setResult.AdvanceOnClick);
        Assert.True(setResult.AdvanceOnTime);
        Assert.Equal(5f, setResult.AdvanceTimeSeconds);
    }

    [Fact]
    public void SetTransition_WithInvalidTransitionName_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetTransition(batch, 1, "ppEffectDoesNotExist");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void SetTransition_WithInvalidSlideIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetTransition(batch, 99, "ppEffectFade");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void GetTransition_OnFreshPresentation_ReturnsSuccessWithDefaultTransition()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetTransition(batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.TransitionName);
        Assert.NotNull(result.DurationSeconds);
        Assert.NotNull(result.AdvanceOnClick);
        Assert.NotNull(result.AdvanceOnTime);
    }
}
