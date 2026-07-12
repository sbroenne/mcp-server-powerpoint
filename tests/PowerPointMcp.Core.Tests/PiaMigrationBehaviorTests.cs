using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real-COM integration tests covering the specific dynamic COM patterns in
/// <see cref="NotesCommands"/> that are most sensitive to the PIA migration.
/// </summary>
/// <remarks>
/// <para>Why these tests exist alongside <see cref="NotesCommandsTests"/>:</para>
/// <list type="bullet">
///   <item><b>Empty-notes placeholder scan</b> — <c>NotesCommands.FindNotesTextShape</c> walks
///   the notes page's Shapes collection dynamically (<c>shape.Type == 14</c> / msoPlaceholder,
///   <c>shape.PlaceholderFormat.Type == 2</c> / ppPlaceholderBody). The existing Notes tests
///   only exercise this path after text has been <em>set</em>. This test exercises the scan on
///   a completely untouched slide, verifying that the placeholder is found and an empty string
///   is returned rather than an exception thrown. If the PIA migration changes the
///   <c>isPlaceholder</c> comparison (e.g., wrong enum value), <c>FindNotesTextShape</c> throws
///   <see cref="InvalidOperationException"/> — this test catches that regression.</item>
///   <item><b>Multi-slide notes access</b> — the existing round-trip test only uses slide 1.
///   This tests slide 2, exercising the <c>ctx.Presentation.Slides[slideIndex].NotesPage</c>
///   dynamic access with a non-first index.</item>
/// </list>
/// <para>No mocking. Requires a live PowerPoint COM instance. Serialized via
/// <c>xunit.runner.json</c> (<c>maxParallelThreads: 1</c>).</para>
/// </remarks>
[Trait("Category", "Integration")]
[Trait("Feature", "Notes")]
public class PiaMigrationBehaviorTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly NotesCommands _notesCommands = new();
    private readonly SlideCommands _slideCommands = new();

    public PiaMigrationBehaviorTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    /// <summary>
    /// Verifies that <c>GetNotesText</c> on a freshly-created slide that has never had any
    /// notes text set returns <c>Success = true</c> with empty (or null) text rather than
    /// throwing an exception.
    /// </summary>
    /// <remarks>
    /// This directly exercises <c>NotesCommands.FindNotesTextShape</c>'s dynamic placeholder
    /// scan: it must locate the body placeholder (<c>shape.Type == 14</c>,
    /// <c>shape.PlaceholderFormat.Type == 2</c>) on an untouched notes page and read its
    /// (empty) text without throwing. If the PIA migration introduces an off-by-one or
    /// wrong-enum comparison in the <c>isPlaceholder</c> check, the scan finds nothing and
    /// throws <see cref="InvalidOperationException"/>, which propagates out of
    /// <c>batch.Execute</c> as an unhandled exception (Rule 1b). This test catches that.
    /// </remarks>
    [Fact]
    public void GetNotesText_OnFreshSlide_WithNoTextSet_ReturnsSuccessWithEmptyText()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        // A fresh presentation has slide 1 with no notes text yet.
        var result = _notesCommands.GetNotesText(batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        // PowerPoint's default body placeholder text is an empty string — not null, not a
        // prompt like "Click to add notes" (that's rendered in the UI but not stored as text).
        Assert.NotNull(result.NotesText);
        Assert.Equal(string.Empty, result.NotesText);
    }

    /// <summary>
    /// Verifies that <c>SetNotesText</c> and <c>GetNotesText</c> work correctly on slide 2
    /// of a two-slide presentation.
    /// </summary>
    /// <remarks>
    /// The existing <see cref="NotesCommandsTests.SetNotesText_ThenGetNotesText_RoundTrips_AndPersistsAfterSave"/>
    /// test only uses slide index 1. This test exercises the <c>ctx.Presentation.Slides[2].NotesPage</c>
    /// access (dynamic dispatch with a non-first index) and confirms the placeholder scan
    /// succeeds on slide 2 as well. A migration bug that breaks the dynamic slide indexer for
    /// non-first slides would cause a <c>COMException</c> or wrong-index failure here.
    /// </remarks>
    [Fact]
    public void SetAndGetNotesText_OnSlide2_WorksCorrectly()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _slideCommands.AddBlank(batch); // presentation now has slides 1 and 2

        var setResult = _notesCommands.SetNotesText(batch, 2, "Speaker notes for slide 2.");

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Null(setResult.ErrorMessage);

        var getResult = _notesCommands.GetNotesText(batch, 2);

        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("Speaker notes for slide 2.", getResult.NotesText);
    }

    /// <summary>
    /// Verifies that <c>GetNotesText</c> on slide 2 with no text set returns an empty string,
    /// not an exception — confirming the dynamic placeholder scan is slide-index-independent.
    /// </summary>
    [Fact]
    public void GetNotesText_OnFreshSlide2_WithNoTextSet_ReturnsSuccessWithEmptyText()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _slideCommands.AddBlank(batch); // 2 slides; slide 2 has no notes

        var result = _notesCommands.GetNotesText(batch, 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.NotNull(result.NotesText);
        Assert.Equal(string.Empty, result.NotesText);
    }

    /// <summary>
    /// Verifies that notes set on slide 1 are independent from notes on slide 2 — no
    /// cross-slide pollution via the dynamic <c>Slides[slideIndex].NotesPage</c> access.
    /// </summary>
    [Fact]
    public void SetNotesText_OnSlide1_DoesNotAffectSlide2_Notes()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _slideCommands.AddBlank(batch); // 2 slides

        _notesCommands.SetNotesText(batch, 1, "Slide 1 notes.");

        var slide2Result = _notesCommands.GetNotesText(batch, 2);

        Assert.True(slide2Result.Success, slide2Result.ErrorMessage);
        // Slide 2 notes must be unaffected — any cross-slide dynamic indexer bug would
        // cause slide 2's notes to read back slide 1's text.
        Assert.Equal(string.Empty, slide2Result.NotesText);
    }
}
