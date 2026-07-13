using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for built-in document property and custom property CRUD, against
/// live PowerPoint COM. No mocking — per Rule 30.
/// </summary>
/// <remarks>
/// Unlike <see cref="PresentationCommandsTests"/> (which genuinely needs a fresh
/// Create()/BeginBatch()/Dispose() cycle per test because it exercises that lifecycle itself),
/// these tests only read/write document properties on an already-open presentation. They share
/// ONE PowerPoint.Application instance across the whole class via
/// <see cref="SharedPresentationFixture"/>, paying PowerPoint's documented ~90-100s post-Quit()
/// process-exit lingering (Ripley's MCP round-trip finding, .squad/decisions.md 2026-07-01) ONCE
/// at fixture disposal instead of once per [Fact]/[Theory] case. Before this split, this file's
/// predecessor (16 test cases, including an 8-case [Theory]) each opened AND closed two separate
/// PowerPoint sessions (one via <c>_commands.Create()</c>, one via
/// <c>PresentationSession.BeginBatch()</c>) — roughly 300s x 16 ≈ 80 minutes of pure COM
/// startup/teardown overhead for tests whose actual assertions run in under a second.
/// </remarks>
[Trait("Category", "Integration")]
[Trait("Feature", "Presentation")]
public class PresentationPropertiesTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _commands = new();

    public PresentationPropertiesTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
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
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetDocumentProperty(batch, propertyName, "Test Value");
        Assert.True(setResult.Success);
        Assert.Null(setResult.ErrorMessage);
        Assert.Equal("Test Value", setResult.PropertyValue);

        var getResult = _commands.GetDocumentProperty(batch, propertyName);
        Assert.True(getResult.Success);
        Assert.Null(getResult.ErrorMessage);
        Assert.Equal("Test Value", getResult.PropertyValue);
    }

    [Fact]
    public void SetDocumentProperty_UnsupportedName_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetDocumentProperty(batch, "NotARealProperty", "value");

        Assert.False(result.Success);
        Assert.Contains("NotARealProperty", result.ErrorMessage);
    }

    [Fact]
    public void GetDocumentProperty_UnsupportedName_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetDocumentProperty(batch, "NotARealProperty");

        Assert.False(result.Success);
        Assert.Contains("NotARealProperty", result.ErrorMessage);
    }

    [Fact]
    public void SetCustomProperty_ThenGetCustomProperty_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
        Assert.True(setResult.Success);
        Assert.Null(setResult.ErrorMessage);

        var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
        Assert.True(getResult.Success);
        Assert.Equal("ABC-123", getResult.PropertyValue);
    }

    [Fact]
    public void SetCustomProperty_CalledTwice_UpdatesExistingValue_DoesNotDuplicate()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
        var secondSet = _commands.SetCustomProperty(batch, "ProjectCode", "XYZ-999");
        Assert.True(secondSet.Success);

        var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
        Assert.Equal("XYZ-999", getResult.PropertyValue);
    }

    [Fact]
    public void GetCustomProperty_NotFound_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetCustomProperty(batch, "DoesNotExist");

        Assert.False(result.Success);
        Assert.Contains("DoesNotExist", result.ErrorMessage);
    }

    [Fact]
    public void RemoveCustomProperty_AfterSet_RemovesIt_SubsequentGetFails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.SetCustomProperty(batch, "ProjectCode", "ABC-123");
        var removeResult = _commands.RemoveCustomProperty(batch, "ProjectCode");
        Assert.True(removeResult.Success);
        Assert.Null(removeResult.ErrorMessage);

        var getResult = _commands.GetCustomProperty(batch, "ProjectCode");
        Assert.False(getResult.Success);
    }

    [Fact]
    public void RemoveCustomProperty_NotFound_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.RemoveCustomProperty(batch, "DoesNotExist");

        Assert.False(result.Success);
        Assert.Contains("DoesNotExist", result.ErrorMessage);
    }
}
