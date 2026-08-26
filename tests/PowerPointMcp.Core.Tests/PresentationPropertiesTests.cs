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
/// process-exit lingering once
/// at fixture disposal instead of once per [Fact]/[Theory] case. Before this split, this file's
/// predecessor (16 test cases, including an 8-case [Theory]) each opened AND closed two separate
/// PowerPoint sessions (one via <c>_commands.Create()</c>, one via
/// <c>PresentationSession.BeginBatch()</c>) — roughly 300s x 16 ≈ 80 minutes of pure COM
/// startup/teardown overhead for tests whose actual assertions run in under a second.
/// </remarks>
[Trait("Category", "Integration")]
[Trait("Feature", "Presentation")]
public partial class PresentationPropertiesTests : IClassFixture<SharedPresentationFixture>
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

    [Fact]
    public void Tags_CrudIsCaseInsensitive_EnumeratesOneBased_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var firstSet = _commands.SetTag(batch, " ReviewState ", "MiXeD Value");
        Assert.True(firstSet.Success, firstSet.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", firstSet.TagName);
        Assert.Equal("MiXeD Value", firstSet.TagValue);
        Assert.Equal(1, firstSet.TagCount);

        var updated = _commands.SetTag(batch, " reviewstate ", "Updated Value");
        Assert.True(updated.Success, updated.ErrorMessage);
        Assert.Equal(1, updated.TagCount);

        var secondSet = _commands.SetTag(batch, "Owner", "Alice");
        Assert.True(secondSet.Success, secondSet.ErrorMessage);
        Assert.Equal(2, secondSet.TagCount);

        var unspacedSet = _commands.SetTag(batch, "ReviewState", "Unspaced Value");
        Assert.True(unspacedSet.Success, unspacedSet.ErrorMessage);
        Assert.Equal(3, unspacedSet.TagCount);

        var get = _commands.GetTag(batch, " ReViEwStAtE ");
        Assert.True(get.Success, get.ErrorMessage);
        Assert.Equal(" REVIEWSTATE ", get.TagName);
        Assert.Equal("Updated Value", get.TagValue);
        Assert.Equal(1, get.TagIndex);

        var unspacedGet = _commands.GetTag(batch, "reviewstate");
        Assert.True(unspacedGet.Success, unspacedGet.ErrorMessage);
        Assert.Equal("REVIEWSTATE", unspacedGet.TagName);
        Assert.Equal("Unspaced Value", unspacedGet.TagValue);
        Assert.Equal(3, unspacedGet.TagIndex);

        var listed = _commands.ListTags(batch);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(3, listed.TagCount);
        Assert.Collection(
            listed.Tags!,
            tag =>
            {
                Assert.Equal(1, tag.TagIndex);
                Assert.Equal(" REVIEWSTATE ", tag.Name);
                Assert.Equal("Updated Value", tag.Value);
            },
            tag =>
            {
                Assert.Equal(2, tag.TagIndex);
                Assert.Equal("OWNER", tag.Name);
                Assert.Equal("Alice", tag.Value);
            },
            tag =>
            {
                Assert.Equal(3, tag.TagIndex);
                Assert.Equal("REVIEWSTATE", tag.Name);
                Assert.Equal("Unspaced Value", tag.Value);
            });

        Assert.True(_commands.Save(batch).Success);
        _fixture.ReopenCurrentPresentation();

        var persisted = _commands.GetTag(batch, " reviewstate ");
        Assert.True(persisted.Success, persisted.ErrorMessage);
        Assert.Equal("Updated Value", persisted.TagValue);

        var deleted = _commands.DeleteTag(batch, " REVIEWSTATE ");
        Assert.True(deleted.Success, deleted.ErrorMessage);
        Assert.Equal(2, deleted.TagCount);

        var missingGet = _commands.GetTag(batch, " reviewstate ");
        Assert.False(missingGet.Success);
        Assert.Contains("REVIEWSTATE", missingGet.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var missingDelete = _commands.DeleteTag(batch, " reviewstate ");
        Assert.False(missingDelete.Success);

        Assert.True(_commands.ListTags(batch).Success);
        Assert.True(_commands.GetTag(batch, "owner").Success);
        Assert.True(_commands.GetTag(batch, "reviewstate").Success);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Tags_WithBlankName_ReturnFailure(string tagName)
    {
        _fixture.CreateFreshPresentation();

        Assert.False(_commands.SetTag(_fixture.Batch, tagName, "value").Success);
        Assert.False(_commands.GetTag(_fixture.Batch, tagName).Success);
        Assert.False(_commands.DeleteTag(_fixture.Batch, tagName).Success);
    }
}
