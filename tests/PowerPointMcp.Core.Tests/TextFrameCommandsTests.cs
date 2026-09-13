using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for text frame commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "TextFrame")]
public class TextFrameCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextFrameCommands _commands = new();

    public TextFrameCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private IPresentationBatch SetUpPresentationWithShape()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        _shapeCommands.AddRectangle(batch, 1, 0f, 0f, 200f, 100f);
        return batch;
    }

    [Fact]
    public void ReplaceText_DoesNotSearchInsertedText()
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, "cat cat").Success);

        var result = _commands.ReplaceText(batch, 1, 1, "cat", "catcat");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.ReplacementCount);
        Assert.Equal("catcat catcat", _commands.GetText(batch, 1, 1).Text);
    }

    [Fact]
    public void SetText_ThenGetText_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();

        var setResult = _commands.SetText(batch, 1, 1, "Quarterly Report");
        Assert.True(setResult.Success);

        var getResult = _commands.GetText(batch, 1, 1);
        Assert.True(getResult.Success);
        Assert.Equal("Quarterly Report", getResult.Text);
    }

    [Fact]
    public void FindText_ReturnsOriginalPositionsWithoutMutation()
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, "cat CAT cat").Success);

        var result = _commands.FindText(batch, 1, 1, "cat");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(3, result.MatchCount);
        Assert.NotNull(result.Matches);
        Assert.Equal([1, 5, 9], result.Matches.Select(match => match.Start));
        Assert.All(result.Matches, match => Assert.Equal(3, match.Length));
        Assert.Equal(["cat", "CAT", "cat"], result.Matches.Select(match => match.Text));
        Assert.Equal("cat CAT cat", _commands.GetText(batch, 1, 1).Text);
    }

    [Theory]
    [InlineData("catcat", "cat", "", false, false, "", 2)]
    [InlineData("aaaaa", "aa", "X", false, false, "XXa", 2)]
    [InlineData("cat cat", "cat", "cat", false, false, "cat cat", 2)]
    [InlineData("Cat cat scatter", "cat", "X", true, true, "Cat X scatter", 1)]
    [InlineData("Cat cat scatter", "cat", "X", false, true, "X X scatter", 2)]
    [InlineData("Cat cat scatter", "cat", "X", true, false, "Cat X sXter", 2)]
    [InlineData("Cat cat scatter", "cat", "X", false, false, "X X sXter", 3)]
    [InlineData("a.b axb a.b", "a.b", "$1", false, false, "$1 axb $1", 2)]
    [InlineData("unchanged", "missing", "X", false, false, "unchanged", 0)]
    [InlineData("cat\ncat", "cat", "X", false, false, "X\nX", 2)]
    [InlineData("caf\u00e9 caf\u00e9", "caf\u00e9", "th\u00e9", true, true, "th\u00e9 th\u00e9", 2)]
    [InlineData("a b c", " ", "", false, false, "abc", 2)]
    [InlineData("", "cat", "X", false, false, "", 0)]
    public void ReplaceText_HandlesLiteralMatches(string text, string findWhat, string replaceWhat,
        bool matchCase, bool wholeWords, string expected, int count)
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, text).Success);

        var result = _commands.ReplaceText(batch, 1, 1, findWhat, replaceWhat, matchCase, wholeWords);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(count, result.ReplacementCount);
        Assert.Equal(expected, _commands.GetText(batch, 1, 1).Text);
    }

    [Theory]
    [InlineData(0, 1, "cat")]
    [InlineData(-1, 1, "cat")]
    [InlineData(2, 1, "cat")]
    [InlineData(1, 0, "cat")]
    [InlineData(1, -1, "cat")]
    [InlineData(1, 2, "cat")]
    [InlineData(1, 1, "")]
    public void ReplaceText_RejectsInvalidInputsWithoutMutation(int slideIndex, int shapeIndex, string findWhat)
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, "cat").Success);

        var result = _commands.ReplaceText(batch, slideIndex, shapeIndex, findWhat, "X");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Null(result.ReplacementCount);
        Assert.Equal("cat", _commands.GetText(batch, 1, 1).Text);

        var findResult = _commands.FindText(batch, slideIndex, shapeIndex, findWhat);
        Assert.False(findResult.Success);
        Assert.False(string.IsNullOrEmpty(findResult.ErrorMessage));
        Assert.Null(findResult.Matches);
    }

    [Theory]
    [InlineData("cat cat", "cat", false, false, 2)]
    [InlineData("aaaaa", "aa", false, false, 2)]
    [InlineData("Cat cat scatter", "cat", true, true, 1)]
    [InlineData("Cat cat scatter", "cat", false, true, 2)]
    [InlineData("a b c", " ", false, false, 2)]
    [InlineData("a.b axb", "a.b", false, false, 1)]
    [InlineData("", "cat", false, false, 0)]
    [InlineData("cat", "missing", false, false, 0)]
    public void FindText_HandlesLiteralMatches(string text, string findWhat, bool matchCase, bool wholeWords, int count)
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, text).Success);

        var result = _commands.FindText(batch, 1, 1, findWhat, matchCase, wholeWords);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(count, result.MatchCount);
        Assert.NotNull(result.Matches);
        Assert.Equal(count, result.Matches.Count);
        Assert.Equal(text, _commands.GetText(batch, 1, 1).Text);
    }

    [Fact]
    public void ReplaceText_PreservesSurroundingFormattingAndOtherShapeAfterReopen()
    {
        var batch = SetUpPresentationWithShape();
        Assert.True(_commands.SetText(batch, 1, 1, "left cat right").Success);
        Assert.True(_shapeCommands.AddRectangle(batch, 1, 250f, 0f, 200f, 100f).Success);
        Assert.True(_commands.SetText(batch, 1, 2, "cat").Success);
        WithTextRange(batch, range =>
        {
            PowerPoint.TextRange? left = null;
            PowerPoint.TextRange? right = null;
            PowerPoint.Font? leftFont = null;
            PowerPoint.Font? rightFont = null;
            try
            {
                left = range.Characters(1, 4);
                right = range.Characters(10, 5);
                leftFont = left.Font;
                rightFont = right.Font;
                leftFont.Size = 11;
                rightFont.Size = 29;
                leftFont.Name = "Arial";
                rightFont.Name = "Georgia";
            }
            finally
            {
                ComUtilities.Release(ref rightFont!);
                ComUtilities.Release(ref leftFont!);
                ComUtilities.Release(ref right!);
                ComUtilities.Release(ref left!);
            }
        });

        var result = _commands.ReplaceText(batch, 1, 1, "cat", "elephant");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ReplacementCount);
        Assert.True(_presentationCommands.Save(batch).Success);
        _fixture.ReopenCurrentPresentation();

        Assert.Equal("left elephant right", _commands.GetText(batch, 1, 1).Text);
        Assert.Equal("cat", _commands.GetText(batch, 1, 2).Text);
        WithTextRange(batch, range =>
        {
            PowerPoint.TextRange? left = null;
            PowerPoint.TextRange? right = null;
            PowerPoint.Font? leftFont = null;
            PowerPoint.Font? rightFont = null;
            try
            {
                left = range.Characters(1, 4);
                right = range.Characters(15, 5);
                leftFont = left.Font;
                rightFont = right.Font;
                Assert.Equal(11f, leftFont.Size);
                Assert.Equal(29f, rightFont.Size);
                Assert.Equal("Arial", leftFont.Name);
                Assert.Equal("Georgia", rightFont.Name);
            }
            finally
            {
                ComUtilities.Release(ref rightFont!);
                ComUtilities.Release(ref leftFont!);
                ComUtilities.Release(ref right!);
                ComUtilities.Release(ref left!);
            }
        });
    }

    [Fact]
    public void FindText_AndReplaceText_RejectShapesWithoutTextFrames()
    {
        var batch = SetUpPresentationWithShape();
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            var image = new ImageCommands().AddPicture(batch, 1, imagePath, 250f, 0f, 100f, 100f);
            Assert.True(image.Success, image.ErrorMessage);

            var found = _commands.FindText(batch, 1, 2, "cat");
            var replaced = _commands.ReplaceText(batch, 1, 2, "cat", "X");
            Assert.False(found.Success);
            Assert.False(replaced.Success);
            Assert.Contains("no text frame", found.ErrorMessage);
            Assert.Contains("no text frame", replaced.ErrorMessage);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    private static void WithTextRange(IPresentationBatch batch, Action<PowerPoint.TextRange> action)
    {
        batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            try
            {
                slides = ctx.Presentation.Slides;
                slide = slides[1];
                shapes = slide.Shapes;
                shape = shapes[1];
                frame = shape.TextFrame;
                range = frame.TextRange;
                action(range);
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref frame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    [Fact]
    public void SetFontSize_And_SetBold_PersistAfterSave()
    {
        var batch = SetUpPresentationWithShape();

        _commands.SetText(batch, 1, 1, "Big Bold Title");
        var sizeResult = _commands.SetFontSize(batch, 1, 1, 40f);
        Assert.True(sizeResult.Success);
        Assert.Equal(40f, sizeResult.FontSize);

        var boldResult = _commands.SetBold(batch, 1, 1, true);
        Assert.True(boldResult.Success);
        Assert.True(boldResult.Bold);

        _presentationCommands.Save(batch);

        _fixture.ReopenCurrentPresentation();
        float size = batch.Execute((ctx, ct) =>
            (float)ctx.Presentation.Slides[1].Shapes[1].TextFrame.TextRange.Font.Size);
        Assert.Equal(40f, size);
    }

    [Fact]
    public void SetFontColor_PacksRgbInPowerPointByteOrder()
    {
        var batch = SetUpPresentationWithShape();

        _commands.SetText(batch, 1, 1, "Red Text");
        var result = _commands.SetFontColor(batch, 1, 1, red: 255, green: 0, blue: 0);

        Assert.True(result.Success);
        Assert.Equal(255, result.ColorRgb); // pure red => 0x0000FF in BGR-packed RGB
    }

    [Fact]
    public void SetFontSize_AndGetFontSize_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Sized Text");

        var setResult = _commands.SetFontSize(batch, 1, 1, 28f);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(28f, setResult.FontSize);

        var getResult = _commands.GetFontSize(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(28f, getResult.FontSize);
    }

    [Fact]
    public void SetBold_AndGetBold_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Bold Text");

        var setResult = _commands.SetBold(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Bold);

        var getResult = _commands.GetBold(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Bold);
    }

    [Fact]
    public void SetFontColor_AndGetFontColor_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Colored Text");

        var setResult = _commands.SetFontColor(batch, 1, 1, red: 255, green: 0, blue: 0);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        var getResult = _commands.GetFontColor(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(255, getResult.ColorRgb); // pure red => 0x0000FF in BGR-packed RGB
    }

    [Fact]
    public void GetBold_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.GetBold(batch, 1, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetText_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.GetText(batch, 1, 99);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetItalic_AndGetItalic_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Italic Text");

        var setResult = _commands.SetItalic(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Italic);

        var getResult = _commands.GetItalic(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Italic);
    }

    [Fact]
    public void SetItalic_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetItalic(batch, 1, 99, true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetUnderline_AndGetUnderline_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Underlined Text");

        var setResult = _commands.SetUnderline(batch, 1, 1, true);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.Underline);

        var getResult = _commands.GetUnderline(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.Underline);
    }

    [Fact]
    public void SetUnderline_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetUnderline(batch, 1, 99, true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetFontName_AndGetFontName_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Styled Text");

        var setResult = _commands.SetFontName(batch, 1, 1, "Georgia");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Georgia", setResult.FontName);

        var getResult = _commands.GetFontName(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("Georgia", getResult.FontName);
    }

    [Fact]
    public void SetFontName_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetFontName(batch, 1, 99, "Georgia");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAlignment_AndGetAlignment_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Centered Text");

        var setResult = _commands.SetAlignment(batch, 1, 1, "ppAlignCenter");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppAlignCenter", setResult.Alignment);

        var getResult = _commands.GetAlignment(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("ppAlignCenter", getResult.Alignment);
    }

    [Fact]
    public void SetAlignment_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetAlignment(batch, 1, 1, "ppAlignDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAlignment_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetAlignment(batch, 1, 99, "ppAlignCenter");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBullet_Enabled_WithCharacter_AndGetBullet_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Bulleted Text");

        var setResult = _commands.SetBullet(batch, 1, 1, enabled: true, character: "-");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.True(setResult.BulletEnabled);
        Assert.Equal("-", setResult.BulletCharacter);

        var getResult = _commands.GetBullet(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.BulletEnabled);
        Assert.Equal("-", getResult.BulletCharacter);
    }

    [Fact]
    public void SetBullet_Disabled_AfterEnabled_TurnsOffBullets()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        _commands.SetBullet(batch, 1, 1, enabled: true, character: "-");
        var disableResult = _commands.SetBullet(batch, 1, 1, enabled: false);
        Assert.True(disableResult.Success, disableResult.ErrorMessage);
        Assert.False(disableResult.BulletEnabled);

        var getResult = _commands.GetBullet(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.BulletEnabled);
    }

    [Fact]
    public void SetBullet_WithMultiCharacterString_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetBullet(batch, 1, 1, enabled: true, character: "ab");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBullet_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetBullet(batch, 1, 99, enabled: true);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAutoSize_AndGetAutoSize_RoundTrips()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Auto-fit Text");

        var setResult = _commands.SetAutoSize(batch, 1, 1, "ppAutoSizeShapeToFitText");
        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("ppAutoSizeShapeToFitText", setResult.AutoSize);

        var getResult = _commands.GetAutoSize(batch, 1, 1);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal("ppAutoSizeShapeToFitText", getResult.AutoSize);
    }

    [Fact]
    public void SetAutoSize_WithUnrecognizedName_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();
        _commands.SetText(batch, 1, 1, "Text");

        var result = _commands.SetAutoSize(batch, 1, 1, "ppAutoSizeDoesNotExist");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetAutoSize_WithInvalidShapeIndex_ReturnsFailure_NotException()
    {
        var batch = SetUpPresentationWithShape();

        var result = _commands.SetAutoSize(batch, 1, 99, "ppAutoSizeNone");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
