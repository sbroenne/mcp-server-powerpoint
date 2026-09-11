using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Master;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide master title/body font and background color operations
/// (<see cref="MasterCommands"/>). No mocking — drives live PowerPoint COM. Shares one
/// PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/>.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Master")]
public class MasterCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly MasterCommands _commands = new();

    public MasterCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GetThemeColors_ReturnsTwelveNamedRgbColorsForSelectedMaster()
    {
        _fixture.CreateFreshPresentation();

        var result = _commands.GetThemeColors(_fixture.Batch, masterIndex: 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(1, result.MasterIndex);
        Assert.False(string.IsNullOrWhiteSpace(result.MasterName));
        Assert.NotNull(result.ThemeColors);
        string[] roles = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2",
            "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"];
        Assert.Equal(roles.Order(), result.ThemeColors.Keys.Order());
        Assert.All(result.ThemeColors.Values, color => Assert.Matches("^#[0-9A-F]{6}$", color));
    }

    [Fact]
    public void GetThemeColors_ReadsDistinctDesignPalettesAndPreservesThemAfterReopen()
    {
        _fixture.CreateFreshPresentation();
        _fixture.Batch.Execute((ctx, ct) =>
        {
            PowerPoint.Designs? designs = null;
            PowerPoint.Design? first = null;
            PowerPoint.Design? second = null;
            try
            {
                designs = ctx.Presentation.Designs;
                first = designs[1];
                second = designs.Add("DistinctPalette");
                SetAccentForPaletteTest(first, 0x913D0B);
                SetAccentForPaletteTest(second, 0x2367C1);
            }
            finally
            {
                ComUtilities.Release(ref second);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref designs);
            }
        });

        var inventory = _commands.ListMasters(_fixture.Batch);
        Assert.True(inventory.Success, inventory.ErrorMessage);
        Assert.Equal(2, inventory.Masters!.Count);
        var firstPalette = _commands.GetThemeColors(_fixture.Batch);
        var secondPalette = _commands.GetThemeColors(_fixture.Batch, 2);
        Assert.True(firstPalette.Success, firstPalette.ErrorMessage);
        Assert.True(secondPalette.Success, secondPalette.ErrorMessage);
        Assert.Equal("#0B3D91", firstPalette.ThemeColors!["Accent1"]);
        Assert.Equal("#C16723", secondPalette.ThemeColors!["Accent1"]);
        Assert.Equal(inventory.Masters[1].MasterName, secondPalette.MasterName);
        Assert.Equal(2, secondPalette.MasterIndex);

        _fixture.Batch.Save();
        _fixture.ReopenCurrentPresentation();
        var reopened = _commands.GetThemeColors(_fixture.Batch, 1);
        Assert.True(reopened.Success, reopened.ErrorMessage);
        Assert.Equal(firstPalette.ThemeColors.OrderBy(entry => entry.Key),
            reopened.ThemeColors!.OrderBy(entry => entry.Key));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(2)]
    public void GetThemeColors_InvalidMaster_ReturnsValidationFailure(int masterIndex)
    {
        _fixture.CreateFreshPresentation();

        var result = _commands.GetThemeColors(_fixture.Batch, masterIndex);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Null(result.ThemeColors);
        Assert.Single(_commands.ListMasters(_fixture.Batch).Masters!);
    }

    private static void SetAccentForPaletteTest(PowerPoint.Design design, int oleColor)
    {
        PowerPoint.Master? master = null;
        Office.OfficeTheme? theme = null;
        Office.ThemeColorScheme? scheme = null;
        Office.ThemeColor? accent = null;
        try
        {
            master = design.SlideMaster;
            theme = master.Theme;
            scheme = theme.ThemeColorScheme;
            accent = scheme.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1);
            accent.RGB = oleColor;
        }
        finally
        {
            ComUtilities.Release(ref accent);
            ComUtilities.Release(ref scheme);
            ComUtilities.Release(ref theme);
            ComUtilities.Release(ref master);
        }
    }

    [Fact]
    public void GetTitleFont_OnFreshPresentation_ReturnsSuccessWithFontDetails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetTitleFont(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.False(string.IsNullOrEmpty(result.FontName));
        Assert.NotNull(result.FontSize);
        Assert.NotNull(result.Bold);
        Assert.NotNull(result.ColorRgb);
    }

    [Fact]
    public void SetTitleFont_ChangesNameSizeBoldAndColor_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetTitleFont(batch, fontName: "Arial", fontSize: 44f, bold: true, red: 200, green: 30, blue: 30);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Arial", setResult.FontName);
        Assert.Equal(44f, setResult.FontSize);
        Assert.True(setResult.Bold);

        var getResult = _commands.GetTitleFont(batch);
        Assert.True(getResult.Success);
        Assert.Equal("Arial", getResult.FontName);
        Assert.Equal(44f, getResult.FontSize);
        Assert.True(getResult.Bold);
        int expectedRgb = 200 + (30 << 8) + (30 << 16);
        Assert.Equal(expectedRgb, getResult.ColorRgb);
    }

    [Fact]
    public void SetTitleFont_WithOnlyFontSize_LeavesOtherPropertiesUnchanged()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var before = _commands.GetTitleFont(batch);
        Assert.True(before.Success);

        var setResult = _commands.SetTitleFont(batch, fontSize: 60f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(60f, setResult.FontSize);
        Assert.Equal(before.FontName, setResult.FontName);
        Assert.Equal(before.Bold, setResult.Bold);
    }

    [Fact]
    public void GetBodyFont_OnFreshPresentation_ReturnsSuccessWithFontDetails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.GetBodyFont(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrEmpty(result.FontName));
    }

    [Fact]
    public void SetBodyFont_ChangesNameAndSize_AndPersists()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetBodyFont(batch, fontName: "Georgia", fontSize: 22f);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal("Georgia", setResult.FontName);
        Assert.Equal(22f, setResult.FontSize);

        var getResult = _commands.GetBodyFont(batch);
        Assert.True(getResult.Success);
        Assert.Equal("Georgia", getResult.FontName);
        Assert.Equal(22f, getResult.FontSize);
    }

    [Fact]
    public void SetAndGetBackgroundColor_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetBackgroundColor(batch, red: 10, green: 20, blue: 30);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        int expectedRgb = 10 + (20 << 8) + (30 << 16);
        Assert.Equal(expectedRgb, setResult.ColorRgb);

        var getResult = _commands.GetBackgroundColor(batch);
        Assert.True(getResult.Success);
        Assert.Equal(expectedRgb, getResult.ColorRgb);
    }

    [Fact]
    public void SetGradientBackground_AndGetGradientBackground_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var setResult = _commands.SetGradientBackground(
            batch,
            red1: 255, green1: 0, blue1: 0,
            red2: 0, green2: 0, blue2: 255,
            gradientStyle: "msoGradientVertical",
            gradientVariant: 2);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        Assert.Equal(255, setResult.ColorRgb);
        Assert.Equal(16711680, setResult.ColorRgb2);
        Assert.Equal("msoGradientVertical", setResult.GradientStyleName);
        Assert.Equal(2, setResult.GradientVariant);

        var getResult = _commands.GetGradientBackground(batch);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(255, getResult.ColorRgb);
        Assert.Equal(16711680, getResult.ColorRgb2);
        Assert.Equal("msoGradientVertical", getResult.GradientStyleName);
        Assert.Equal(2, getResult.GradientVariant);
    }

    [Fact]
    public void SetGradientBackground_WithUnrecognizedStyleName_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.SetGradientBackground(
            batch,
            red1: 255, green1: 0, blue1: 0,
            red2: 0, green2: 0, blue2: 255,
            gradientStyle: "msoGradientNotARealStyle");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetGradientBackground_WhenBackgroundIsSolid_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        _commands.SetBackgroundColor(batch, red: 255, green: 0, blue: 0);

        var result = _commands.GetGradientBackground(batch);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void ListMasters_ReturnsMastersAndLayouts()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.ListMasters(batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.Masters);
        Assert.NotEmpty(result.Masters);

        var firstMaster = Assert.Single(result.Masters);
        Assert.NotNull(firstMaster.MasterName);
        Assert.True(firstMaster.MasterIndex >= 1);
        Assert.NotNull(firstMaster.Layouts);
        Assert.NotEmpty(firstMaster.Layouts);
    }

    [Fact]
    public void DeleteMaster_RemovesUnusedMasterAfterTemplateApply()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        batch.Execute((ctx, ct) =>
        {
            PowerPoint.Designs designs = ctx.Presentation.Designs;
            PowerPoint.Design extraDesign = designs.Add("PptMcpTestExtraMaster");
            return 0;
        });

        var listResult = _commands.ListMasters(batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        Assert.True(listResult.Masters!.Count >= 2);

        var candidateMaster = listResult.Masters!.First(m => m.MasterName == "PptMcpTestExtraMaster");
        var deleteResult = _commands.DeleteMaster(batch, candidateMaster.MasterIndex);

        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);

        var afterDelete = _commands.ListMasters(batch);
        Assert.True(afterDelete.Success, afterDelete.ErrorMessage);
        Assert.DoesNotContain(afterDelete.Masters!, m => m.MasterName == candidateMaster.MasterName);
        Assert.Equal(listResult.Masters!.Count - 1, afterDelete.Masters!.Count);
    }
}
