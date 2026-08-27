using Sbroenne.PowerPointMcp.Core.Media;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for media commands against live PowerPoint COM. Audio and video
/// fixtures are repository-owned synthetic media with no third-party authored content. An
/// unavailable host video decoder fails explicitly during insertion rather than skipping coverage.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Media")]
public sealed class MediaCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PresentationCommands _presentationCommands = new();
    private readonly MediaCommands _commands = new();

    public MediaCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddMedia_EmbeddedAudio_SurvivesSourceDeletionAndSaveReopen()
    {
        _fixture.CreateFreshPresentation();
        string mediaPath = CreateWaveFixture();
        var addResult = _commands.AddMedia(
            _fixture.Batch, 1, mediaPath, false, true, 10f, 10f, 120f, 60f);

        Assert.True(addResult.Success, addResult.ErrorMessage);
        Assert.Null(addResult.ErrorMessage);
        Assert.Equal(1, addResult.ShapeIndex);
        Assert.Equal(1, addResult.ShapeCount);
        Assert.False(addResult.LinkToFile);
        Assert.True(addResult.SaveWithDocument);

        File.Delete(mediaPath);
        Assert.False(File.Exists(mediaPath));

        _presentationCommands.Save(_fixture.Batch);
        _fixture.ReopenCurrentPresentation();

        var info = _commands.GetMediaInfo(_fixture.Batch, 1, 1);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.Equal("ppMediaTypeSound", info.MediaTypeName);
        Assert.Equal(1, info.ShapeIndex);
        Assert.Equal(1, info.ShapeCount);
    }

    [Fact]
    public void AddMedia_EmbeddedVideo_UsesSyntheticMp4AndPersists()
    {
        _fixture.CreateFreshPresentation();
        string mediaPath = CreateVideoFixture();
        try
        {
            var addResult = _commands.AddMedia(
                _fixture.Batch, 1, mediaPath, false, true, 20f, 20f, 160f, 90f);

            Assert.True(addResult.Success, addResult.ErrorMessage);
            Assert.Equal(1, addResult.ShapeIndex);
            Assert.Equal(1, addResult.ShapeCount);

            _presentationCommands.Save(_fixture.Batch);
            _fixture.ReopenCurrentPresentation();

            var info = _commands.GetMediaInfo(_fixture.Batch, 1, 1);
            Assert.True(info.Success, info.ErrorMessage);
            Assert.Equal("ppMediaTypeMovie", info.MediaTypeName);
            Assert.Equal(1, info.ShapeIndex);
            Assert.Equal(1, info.ShapeCount);
        }
        finally
        {
            File.Delete(mediaPath);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddMedia_LinkedAudioAndVideo_PersistAndReleaseSourceFile(bool video)
    {
        _fixture.CreateFreshPresentation();
        string mediaPath = video ? CreateVideoFixture() : CreateWaveFixture();
        var addResult = _commands.AddMedia(
            _fixture.Batch, 1, mediaPath, true, false, 5f, 5f, 100f, 75f);

        Assert.True(addResult.Success, addResult.ErrorMessage);
        Assert.True(addResult.LinkToFile);
        Assert.False(addResult.SaveWithDocument);
        Assert.Equal(Path.GetFullPath(mediaPath), addResult.SourcePath);

        _presentationCommands.Save(_fixture.Batch);
        _fixture.ReopenCurrentPresentation();

        var info = _commands.GetMediaInfo(_fixture.Batch, 1, 1);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.Equal(video ? "ppMediaTypeMovie" : "ppMediaTypeSound", info.MediaTypeName);

        File.Delete(mediaPath);
        Assert.False(File.Exists(mediaPath));
    }

    [Fact]
    public void AddMedia_WithMissingPath_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var result = _commands.AddMedia(
            _fixture.Batch, 1, "C:\\does\\not\\exist.wav", false, true, 0f, 0f, 100f, 100f);

        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void AddMedia_WithAmbiguousStorageCombination_ReturnsFailure(
        bool linkToFile,
        bool saveWithDocument)
    {
        _fixture.CreateFreshPresentation();
        string mediaPath = CreateWaveFixture();
        try
        {
            var result = _commands.AddMedia(
                _fixture.Batch, 1, mediaPath, linkToFile, saveWithDocument,
                0f, 0f, 100f, 100f);

            Assert.False(result.Success);
            Assert.Contains("combination", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(mediaPath);
        }
    }

    [Fact]
    public void AddMedia_WithMalformedPath_ReturnsFailure()
    {
        var result = _commands.AddMedia(
            _fixture.Batch, 1, "invalid\0media.wav", false, true,
            0f, 0f, 100f, 100f);

        Assert.False(result.Success);
        Assert.Contains("Invalid media path", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void GetMediaInfo_WithInvalidIndexesOrNonMediaShape_ReturnsFailure()
    {
        _fixture.CreateFreshPresentation();
        var shapeCommands = new Core.Shape.ShapeCommands();

        var invalidSlide = _commands.GetMediaInfo(_fixture.Batch, 0, 1);
        Assert.False(invalidSlide.Success);

        var invalidShape = _commands.GetMediaInfo(_fixture.Batch, 1, 1);
        Assert.False(invalidShape.Success);

        var rectangle = shapeCommands.AddRectangle(_fixture.Batch, 1, 0f, 0f, 100f, 100f);
        Assert.True(rectangle.Success, rectangle.ErrorMessage);

        var nonMedia = _commands.GetMediaInfo(_fixture.Batch, 1, 1);
        Assert.False(nonMedia.Success);
        Assert.Contains("not a media", nonMedia.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    private static string CreateWaveFixture()
    {
        string path = CreateFixturePath(".wav");

        // 100 ms of a synthetic 440 Hz tone: 8 kHz, 8-bit unsigned mono PCM.
        // Generated once with PowerShell/.NET by writing the standard RIFF/WAVE header followed by
        // 800 samples computed as 128 + sin(2*pi*440*i/8000)*24. No runtime encoder or external
        // asset is used. These deterministic bytes contain no third-party authored media and are
        // repository-owned under the MIT license.
        const string base64Wav =
            "UklGRkQDAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YSADAACAiI+VmJiVkImBeXFsaWhrcHZ+h46Ul5iWkYqCenNtaWhqb3V9hY2Tl5iWkoyEfHRuaWhpbnR8hIySlpiXk42FfXVvamhpbXN6goqRlpiXlI6HfnZwa2hpbHF5gYmQlZiYlY+IgHhxa2hoa3B3f4ePlJeYlZCKgnlybGloam92foaNk5eYlpGLg3tzbWloam50fISMkpeYl5KMhHx0bmpoaW1ze4OLkZaYl5ONhn52b2poaWxyeYKKkJWYl5SPh393cGtoaGtxeICIj5WYmJWQiYF5cWxpaGtwdn6HjpSXmJaRioJ6c21paGpvdX2FjZOXmJaSjIR8dG5paGludHyEjJKWmJeTjYV9dW9qaGltc3qCipGWmJeUjod+dnBraGlscXmBiZCVmJiVj4iAeHFraGhrcHd/h4+Ul5iVkIqCeXJsaWhqb3Z+ho2Tl5iWkYuDe3NtaWhqbnR8hIySl5iXkoyEfHRuamhpbXN7g4uRlpiXk42GfnZvamhpbHJ5goqQlZiXlI+Hf3dwa2hoa3F4gIiPlZiYlZCJgXlxbGloa3B2foeOlJeYlpGKgnpzbWloam91fYWNk5eYlpKMhHx0bmloaW50fISMkpaYl5ONhX11b2poaW1zeoKKkZaYl5SOh352cGtoaWxxeYGJkJWYmJWPiIB4cWtoaGtwd3+Hj5SXmJWQioJ5cmxpaGpvdn6GjZOXmJaRi4N7c21paGpudHyEjJKXmJeSjIR8dG5qaGltc3uDi5GWmJeTjYZ+dm9qaGlscnmCipCVmJeUj4d/d3BraGhrcXiAiI+VmJiVkImBeXFsaWhrcHZ+h46Ul5iWkYqCenNtaWhqb3V9hY2Tl5iWkoyEfHRuaWhpbnR8hIySlpiXk42FfXVvamhpbXN6goqRlpiXlI6HfnZwa2hpbHF5gYmQlZiYlY+IgHhxa2hoa3B3f4ePlJeYlZCKgnlybGloam92foaNk5eYlpGLg3tzbWloam50fISMkpeYl5KMhHx0bmpoaW1ze4OLkZaYl5ONhn52b2poaWxyeYKKkJWYl5SPh393cGtoaGtxeA==";

        File.WriteAllBytes(path, Convert.FromBase64String(base64Wav));
        return path;
    }

    private static string CreateFixturePath(string extension)
    {
        string directory = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests");
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, $"pptmcp-owned-media-{Guid.NewGuid():N}{extension}");
    }

    private static string CreateVideoFixture()
    {
        string path = CreateFixturePath(".mp4");

        // One second of black video: H.264 Constrained Baseline, 160x90, yuv420p, 10 fps, no audio.
        // Generated solely from FFmpeg's color source with the FFmpeg 8.1 BtbN LGPL ARM64 build
        // (https://github.com/BtbN/FFmpeg-Builds) and libopenh264. Command: ffmpeg -f lavfi -i
        // "color=c=black:size=160x90:rate=10:duration=1" -c:v libopenh264 -pix_fmt yuv420p
        // -movflags +faststart -an output.mp4. The generator is not redistributed. These synthetic
        // bytes contain no third-party authored media and are repository-owned under the MIT license.
        const string base64Mp4 =
            "AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAOAbW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAA+gAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAm10cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAA+gAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAKAAAABaAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAPoAAAAAAABAAAAAAHlbWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAAAoAAAAKABVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABkG1pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAAVBzdGJsAAAAsHN0c2QAAAAAAAAAAQAAAKBhdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAKAAWgBIAAAASAAAAAAAAAABGUxhdmM2Mi4yOC4xMDIgbGlib3BlbmgyNjQAAAAAAAAAGP//AAAAJmF2Y0MBQsAU/+EAD2dCwBSMaCjXkwEB4RCNQAEABGjOPIAAAAAQcGFzcAAAAAEAAAABAAAAFGJ0cnQAAAAAAAAFiAAABYgAAAAYc3R0cwAAAAAAAAABAAAACgAABAAAAAAUc3RzcwAAAAAAAAABAAAAAQAAABxzdHNjAAAAAAAAAAEAAAABAAAACgAAAAEAAAA8c3RzegAAAAAAAAAAAAAACgAAAEUAAAAMAAAADAAAAAwAAAAMAAAADAAAAAwAAAAMAAAADAAAAAwAAAAUc3RjbwAAAAAAAAABAAADsAAAAJ91ZHRhAAAAl21ldGEAAAAAAAAAIWhkbHIAAAAAAAAAAG1kaXJhcHBsAAAAAAAAAAAAAAAAamlsc3QAAAA9qW5hbQAAADVkYXRhAAAAAQAAAABQb3dlclBvaW50TWNwIHN5bnRoZXRpYyBtZWRpYSBmaXh0dXJlAAAAJal0b28AAAAdZGF0YQAAAAEAAAAATGF2ZjYyLjEyLjEwMgAAAAhmcmVlAAAAuW1kYXQAAABBZbgABAnkxQABGfk5OTk5OTk5OTrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr/+P8EBAaTZffffffffffgAAAAIYeAAfkCeD2AAAAAIYeAAvkD+D2AAAAAIYeAA/kBXg9gAAAAIYeABPkBvg9gAAAAIYeABfkB3g9gAAAAIYeABvkB3g9gAAAAIYeAB/kB3g9gAAAAIYeACPkB3g9gAAAAIYeACfkB3g9g=";

        File.WriteAllBytes(path, Convert.FromBase64String(base64Mp4));
        return path;
    }
}
