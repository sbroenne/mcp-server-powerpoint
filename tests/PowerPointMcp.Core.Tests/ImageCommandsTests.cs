using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for image commands against live PowerPoint COM. No mocking.
/// Shares one PowerPoint.Application instance across all [Fact]s in this class via
/// <see cref="SharedPresentationFixture"/> — each test still gets its own freshly-created
/// presentation file for isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Image")]
public class ImageCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly ImageCommands _commands = new();

    public ImageCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddPicture_IncreasesShapeCount_AndPersistsAfterSave()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            var result = _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.ShapeIndex);
            Assert.Equal(1, result.ShapeCount);

            _presentationCommands.Save(batch);

            _fixture.ReopenCurrentPresentation();
            int shapeCount = batch.Execute((ctx, ct) => ctx.Presentation.Slides[1].Shapes.Count);
            Assert.Equal(1, shapeCount);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void AddPicture_WithMissingFile_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;

        var result = _commands.AddPicture(batch, 1, "C:\\does\\not\\exist.png", 0f, 0f, 10f, 10f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetBrightnessContrast_AndGetBrightnessContrast_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var setResult = _commands.SetBrightnessContrast(batch, 1, 1, 0.6f, 0.7f);
            Assert.True(setResult.Success, setResult.ErrorMessage);
            Assert.Equal(0.6f, setResult.Brightness);
            Assert.Equal(0.7f, setResult.Contrast);

            var getResult = _commands.GetBrightnessContrast(batch, 1, 1);
            Assert.True(getResult.Success, getResult.ErrorMessage);
            Assert.Equal(0.6f, getResult.Brightness);
            Assert.Equal(0.7f, getResult.Contrast);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void SetRecolor_AndGetRecolor_RoundTrips()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var setResult = _commands.SetRecolor(batch, 1, 1, "msoPictureGrayscale");
            Assert.True(setResult.Success, setResult.ErrorMessage);
            Assert.Equal("msoPictureGrayscale", setResult.ColorTypeName);

            var getResult = _commands.GetRecolor(batch, 1, 1);
            Assert.True(getResult.Success, getResult.ErrorMessage);
            Assert.Equal("msoPictureGrayscale", getResult.ColorTypeName);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void SetRecolor_WithUnrecognizedColorTypeName_Fails()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var result = _commands.SetRecolor(batch, 1, 1, "msoPictureNotARealType");

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    // ─── SetCrop / GetCrop — permanent red tests (TDD: red before Parker implements) ─
    //
    // CONTRACT (to be implemented by Parker in IImageCommands / ImageCommands):
    //
    //   ImageOperationResult SetCrop(
    //       IPresentationBatch batch,
    //       int slideIndex,
    //       int shapeIndex,
    //       float cropLeft,   // points, left edge; L-T-R-B ordering throughout
    //       float cropTop,    // points, top edge
    //       float cropRight,  // points, right edge
    //       float cropBottom) // points, bottom edge
    //
    //   ImageOperationResult GetCrop(
    //       IPresentationBatch batch,
    //       int slideIndex,
    //       int shapeIndex)
    //
    // ImageOperationResult additions required:
    //   public float? CropLeft   { get; init; }
    //   public float? CropTop    { get; init; }
    //   public float? CropRight  { get; init; }
    //   public float? CropBottom { get; init; }
    //
    // EMPIRICAL BASIS (from passing COM investigation tests below):
    //   • CropLeft/Top/Right/Bottom are in POINTS, not fractions or percentages.
    //   • Round-trip precision: < 0.001 pt; tolerance for tests: 0.01 pt.
    //   • The four sides are independent; setting one does not affect the others.
    //   • Default is 0.0 for all four sides immediately after AddPicture.
    //   • A properly-sized source image (≥ 50×50 px BMP) is required for reliable
    //     round-trip; the 1×1-pixel PNG helper overflows EMU arithmetic (see
    //     CreateProperSizeTestImageFile doc comment below).
    //   • Negative crop values are valid COM inputs; they expand the visible area
    //     beyond the natural image boundary. No range validation needed.
    //
    // THESE TESTS CURRENTLY FAIL TO COMPILE (CS1061) because SetCrop, GetCrop, and
    // CropLeft/Top/Right/Bottom do not yet exist. That is the intentional TDD red state.

    /// <summary>
    /// Primary round-trip: set all four sides with distinct values, assert SetCrop
    /// result carries applied values, then verify via GetCrop.
    /// </summary>
    [Fact]
    public void SetCrop_AndGetCrop_RoundTrip_AllFourSides()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile(); // 50×50 BMP required
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            // L=5, T=3, R=7, B=2 — distinct values confirm no cross-side pollution.
            var setResult = _commands.SetCrop(batch, 1, 1,
                cropLeft: 5f, cropTop: 3f, cropRight: 7f, cropBottom: 2f);

            Assert.True(setResult.Success, setResult.ErrorMessage);
            Assert.Null(setResult.ErrorMessage);
            Assert.NotNull(setResult.CropLeft);
            Assert.InRange(setResult.CropLeft!.Value,  5f - 0.01f, 5f + 0.01f);
            Assert.InRange(setResult.CropTop!.Value,   3f - 0.01f, 3f + 0.01f);
            Assert.InRange(setResult.CropRight!.Value, 7f - 0.01f, 7f + 0.01f);
            Assert.InRange(setResult.CropBottom!.Value,2f - 0.01f, 2f + 0.01f);

            var getResult = _commands.GetCrop(batch, 1, 1);

            Assert.True(getResult.Success, getResult.ErrorMessage);
            Assert.Null(getResult.ErrorMessage);
            Assert.NotNull(getResult.CropLeft);
            Assert.InRange(getResult.CropLeft!.Value,  5f - 0.01f, 5f + 0.01f);
            Assert.InRange(getResult.CropTop!.Value,   3f - 0.01f, 3f + 0.01f);
            Assert.InRange(getResult.CropRight!.Value, 7f - 0.01f, 7f + 0.01f);
            Assert.InRange(getResult.CropBottom!.Value,2f - 0.01f, 2f + 0.01f);
        }
        finally { File.Delete(imagePath); }
    }

    /// <summary>
    /// Verifies that GetCrop returns all-zero crop values on a
    /// freshly-added picture (no crop applied yet) — matching the COM invariant.
    /// </summary>
    [Fact]
    public void GetCrop_OnFreshlyAddedPicture_ReturnsAllZeroes()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            var result = _commands.GetCrop(batch, 1, 1);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.NotNull(result.CropLeft);
            Assert.Equal(0f, result.CropLeft!.Value);
            Assert.Equal(0f, result.CropTop!.Value);
            Assert.Equal(0f, result.CropRight!.Value);
            Assert.Equal(0f, result.CropBottom!.Value);
        }
        finally { File.Delete(imagePath); }
    }

    /// <summary>
    /// Verifies that crop values survive Save → reopen: the
    /// values stored by SetCrop are faithfully read back by GetCrop after the
    /// presentation is closed and re-opened from disk.
    /// </summary>
    [Fact]
    public void SetCrop_Persistence_AfterSaveAndReopen()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            var setResult = _commands.SetCrop(batch, 1, 1,
                cropLeft: 8f, cropTop: 4f, cropRight: 6f, cropBottom: 3f);
            Assert.True(setResult.Success, setResult.ErrorMessage);

            _presentationCommands.Save(batch);
            _fixture.ReopenCurrentPresentation();

            var getResult = _commands.GetCrop(batch, 1, 1);

            Assert.True(getResult.Success, getResult.ErrorMessage);
            Assert.NotNull(getResult.CropLeft);
            Assert.InRange(getResult.CropLeft!.Value,  8f - 0.01f, 8f + 0.01f);
            Assert.InRange(getResult.CropTop!.Value,   4f - 0.01f, 4f + 0.01f);
            Assert.InRange(getResult.CropRight!.Value, 6f - 0.01f, 6f + 0.01f);
            Assert.InRange(getResult.CropBottom!.Value,3f - 0.01f, 3f + 0.01f);
        }
        finally { File.Delete(imagePath); }
    }

    /// <summary>
    /// SetCrop on a non-picture shape must return Success=false
    /// with a non-empty error message — not throw an unhandled COM exception.
    /// </summary>
    [Fact]
    public void SetCrop_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch); // shape 1 is a text box

        var result = _commands.SetCrop(batch, 1, 1,
            cropLeft: 5f, cropTop: 3f, cropRight: 7f, cropBottom: 2f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    /// <summary>
    /// GetCrop on a non-picture shape must return Success=false
    /// with a non-empty error message — not throw an unhandled COM exception.
    /// </summary>
    [Fact]
    public void GetCrop_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch);

        var result = _commands.GetCrop(batch, 1, 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    /// <summary>
    /// SetCrop with an out-of-range slide index must return
    /// Success=false (validated before any COM call, consistent with other Image methods).
    /// </summary>
    [Fact]
    public void SetCrop_WithInvalidSlideIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            var result = _commands.SetCrop(batch, 99, 1, 5f, 3f, 7f, 2f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally { File.Delete(imagePath); }
    }

    /// <summary>
    /// SetCrop with an out-of-range shape index must return
    /// Success=false (validated before any COM call).
    /// </summary>
    [Fact]
    public void SetCrop_WithInvalidShapeIndex_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            var result = _commands.SetCrop(batch, 1, 99, 5f, 3f, 7f, 2f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally { File.Delete(imagePath); }
    }

    // ─── Helpers ──────────────────────────────────────────────────────────────────

    /// <summary>
    /// Adds a text box to slide 1 of the current presentation so tests can exercise
    /// PictureFormat operations on a shape that is NOT a picture. Uses direct COM
    /// dispatch (msoTextOrientationHorizontal = 1) — does not go through ShapeCommands.
    /// </summary>
    private static void AddNonPictureShape(IPresentationBatch batch) =>
        batch.Execute((ctx, ct) =>
        {
            dynamic slide = ctx.Presentation.Slides[1];
            slide.Shapes.AddTextbox(1, 10.0f, 10.0f, 200.0f, 50.0f); // 1 = msoTextOrientationHorizontal
        });

    /// <summary>
    /// Creates a 50×50-pixel BMP at 96 DPI (≈37.5 pt natural size), which is large
    /// enough for PowerPoint to correctly honour the explicit Width/Height passed to
    /// AddPicture and produce stable PictureFormat crop round-trip behaviour.
    /// <para>
    /// Background: the 1×1-pixel PNG produced by
    /// <see cref="CoreTestHelper.CreateUniqueTestImageFile"/> is too small — PowerPoint
    /// internally rescales the shape when its CropLeft is set (empirically: setting
    /// CropLeft=5 on a 100-pt-display shape with a 1×1-pixel source reads back as −2.85
    /// and expands the shape to ≈480 pt). Any test asserting a crop round-trip must use
    /// a properly-sized source image; this factory provides one.
    /// </para>
    /// </summary>
    private static string CreateProperSizeTestImageFile()
    {
        string dir = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, $"pptmcp-test-50px-{Guid.NewGuid():N}.bmp");

        const int W = 50, H = 50;
        int rowStride = (W * 3 + 3) / 4 * 4;   // BMP rows are 4-byte aligned
        byte[] bmp = new byte[54 + rowStride * H];

        // BITMAPFILEHEADER (14 bytes)
        bmp[0] = 0x42; bmp[1] = 0x4D;                           // "BM" signature
        BitConverter.GetBytes(bmp.Length).CopyTo(bmp, 2);        // total file size
        BitConverter.GetBytes(54).CopyTo(bmp, 10);               // pixel data offset

        // BITMAPINFOHEADER (40 bytes at offset 14)
        BitConverter.GetBytes(40).CopyTo(bmp, 14);               // header size
        BitConverter.GetBytes(W).CopyTo(bmp, 18);                // width in pixels
        BitConverter.GetBytes(-H).CopyTo(bmp, 22);               // height — negative = top-down rows
        bmp[26] = 1;                                              // colour planes = 1
        bmp[28] = 24;                                             // bits per pixel = 24 (RGB)
        BitConverter.GetBytes(3780).CopyTo(bmp, 38);             // X pixels per metre ≈ 96 DPI
        BitConverter.GetBytes(3780).CopyTo(bmp, 42);             // Y pixels per metre ≈ 96 DPI

        // Pixel data: solid blue, stored as B G R (BMP byte order)
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                int i = 54 + y * rowStride + x * 3;
                bmp[i] = 180; bmp[i + 1] = 100; bmp[i + 2] = 0;
            }

        File.WriteAllBytes(path, bmp);
        return path;
    }

    // ─── Non-picture shape: PictureFormat operations must fail gracefully ─────────
    // ImageCommands validates the shape type before accessing PictureFormat.
    // These tests verify that operations on non-picture shapes return Success=false
    // with appropriate error messages instead of throwing exceptions.

    [Fact]
    public void SetBrightnessContrast_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch);  // shape 1 is a text box, not a picture

        var result = _commands.SetBrightnessContrast(batch, 1, 1, 0.6f, 0.7f);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetBrightnessContrast_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch);

        var result = _commands.GetBrightnessContrast(batch, 1, 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void SetRecolor_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch);

        var result = _commands.SetRecolor(batch, 1, 1, "msoPictureGrayscale");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    [Fact]
    public void GetRecolor_OnNonPictureShape_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        AddNonPictureShape(batch);

        var result = _commands.GetRecolor(batch, 1, 1);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
    }

    // ─── Brightness / contrast out-of-range validation ─────────────────────────
    // IImageCommands documents brightness and contrast as floats in [0, 1].
    // These tests verify that out-of-range values are validated and return Success=false
    // with appropriate error messages instead of passing invalid values to COM.

    [Fact]
    public void SetBrightnessContrast_WithBrightnessAboveOne_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var result = _commands.SetBrightnessContrast(batch, 1, 1, 1.5f, 0.5f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void SetBrightnessContrast_WithContrastBelowZero_ReturnsFailure_NotException()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CoreTestHelper.CreateUniqueTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 10f, 10f, 100f, 100f);

            var result = _commands.SetBrightnessContrast(batch, 1, 1, 0.5f, -0.1f);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    // ─── PictureFormat crop: COM contract documentation ────────────────────────────
    // These tests access PictureFormat.Crop* directly via batch.Execute because
    // IImageCommands does not yet expose SetCrop/GetCrop operations.  They document
    // the exact COM semantics (units, round-trip precision, independence of the four
    // sides) that any future SetCrop/GetCrop implementation must satisfy.
    //
    // KEY FINDING from COM investigation:
    //   • Crop values are in POINTS (absolute, not percentages or fractions).
    //   • Round-trip precision: < 0.001 pt error for values in 0–50 pt range.
    //   • The four sides are independent; setting one does not affect the others.
    //   • Source-image size matters: for the 1×1-pixel CoreTestHelper PNG, setting
    //     CropLeft=5 reads back as −2.85 and the shape expands to ≈480 pt — an EMU
    //     arithmetic overflow artefact of the near-zero natural image size. A properly-
    //     sized source (≥ a few pixels) round-trips cleanly at the tolerances asserted
    //     below. Any SetCrop/GetCrop implementation MUST use or document this constraint.
    //   • Default ColorType = 1 (msoPictureAutomatic); default Brightness = Contrast = 0.5.

    [Fact]
    public void PictureFormat_CropValues_AreAllZero_AfterAddPicture()
    {
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            var (cropLeft, cropTop, cropRight, cropBottom) = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic pf = slide.Shapes[1].PictureFormat;
                return ((float)pf.CropLeft, (float)pf.CropTop,
                        (float)pf.CropRight, (float)pf.CropBottom);
            });

            Assert.Equal(0f, cropLeft);
            Assert.Equal(0f, cropTop);
            Assert.Equal(0f, cropRight);
            Assert.Equal(0f, cropBottom);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void PictureFormat_CropLeft_SetAndReadBack_RoundTrips_WithinSubPointTolerance()
    {
        // Crop values are in points. Observed round-trip precision: < 0.001 pt
        // (e.g. set 10 pt → read back 9.999646 pt; set 5 pt → 5.000198 pt).
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic shape = slide.Shapes[1];
                shape.PictureFormat.CropLeft = 10f;
            });
            float cropLeft = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic shape = slide.Shapes[1];
                return (float)shape.PictureFormat.CropLeft;
            });

            Assert.InRange(cropLeft, 10f - 0.01f, 10f + 0.01f);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }

    [Fact]
    public void PictureFormat_AllFourCropSides_SetAndReadBack_Independently_WithinTolerance()
    {
        // Verifies that each crop side is stored independently (setting L=5, T=3, R=7, B=2
        // does not bleed into other sides) and that all four round-trip within 0.01 pt.
        _fixture.CreateFreshPresentation();
        var batch = _fixture.Batch;
        string imagePath = CreateProperSizeTestImageFile();
        try
        {
            _commands.AddPicture(batch, 1, imagePath, 0f, 0f, 100f, 100f);

            batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic pf = slide.Shapes[1].PictureFormat;
                pf.CropLeft = 5f; pf.CropTop = 3f; pf.CropRight = 7f; pf.CropBottom = 2f;
            });
            var (cl, ctop, cr, cb) = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic pf = slide.Shapes[1].PictureFormat;
                return ((float)pf.CropLeft, (float)pf.CropTop,
                        (float)pf.CropRight, (float)pf.CropBottom);
            });

            Assert.InRange(cl,   5f - 0.01f, 5f + 0.01f);
            Assert.InRange(ctop, 3f - 0.01f, 3f + 0.01f);
            Assert.InRange(cr,   7f - 0.01f, 7f + 0.01f);
            Assert.InRange(cb,   2f - 0.01f, 2f + 0.01f);
        }
        finally
        {
            File.Delete(imagePath);
        }
    }
}
