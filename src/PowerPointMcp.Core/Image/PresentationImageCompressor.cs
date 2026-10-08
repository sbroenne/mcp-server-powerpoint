using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Xml.Linq;

namespace Sbroenne.PowerPointMcp.Core.Image;

internal sealed record PresentationCompressionResult(
    string Resolution,
    int CompressedPictureCount,
    IReadOnlyList<string> SkippedPictures,
    long OriginalImageBytes,
    long CompressedImageBytes);

[SupportedOSPlatform("windows")]
internal static class PresentationImageCompressor
{
    private static readonly XNamespace PresentationNamespace =
        "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    private const double EmusPerInch = 914400d;

    public static PresentationCompressionResult Compress(
        string presentationPath,
        IReadOnlyDictionary<int, HashSet<int>> selectedPictures,
        string resolution,
        int? targetPpi,
        bool deleteCroppedAreas,
        CancellationToken cancellationToken)
    {
        string sourceCopyPath = Path.Combine(
            Path.GetDirectoryName(presentationPath)!,
            $".{Path.GetFileNameWithoutExtension(presentationPath)}-{Guid.NewGuid():N}.source.zip");
        try
        {
            File.Copy(presentationPath, sourceCopyPath);
            return CompressCopy(
                sourceCopyPath,
                presentationPath,
                selectedPictures,
                resolution,
                targetPpi,
                deleteCroppedAreas,
                cancellationToken);
        }
        finally
        {
            if (File.Exists(sourceCopyPath))
            {
                File.Delete(sourceCopyPath);
            }
        }
    }

    private static PresentationCompressionResult CompressCopy(
        string sourceCopyPath,
        string presentationPath,
        IReadOnlyDictionary<int, HashSet<int>> selectedPictures,
        string resolution,
        int? targetPpi,
        bool deleteCroppedAreas,
        CancellationToken cancellationToken)
    {
        using ZipArchive source = ZipFile.OpenRead(sourceCopyPath);
        List<string> slidePaths = GetSlidePaths(source);
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var addedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var changedImageParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var skipped = new List<string>();
        long originalBytes = 0;
        long compressedBytes = 0;
        int compressedCount = 0;

        foreach ((int slideIndex, HashSet<int> shapeIds) in selectedPictures)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (slideIndex < 1 || slideIndex > slidePaths.Count)
            {
                throw new InvalidDataException($"Presentation package has no slide {slideIndex}.");
            }

            string slidePath = slidePaths[slideIndex - 1];
            string relationshipsPath = GetRelationshipsPath(slidePath);
            ZipArchiveEntry slideEntry = RequireEntry(source, slidePath);
            ZipArchiveEntry relationshipsEntry = RequireEntry(source, relationshipsPath);
            XDocument slideDocument = ReadXml(slideEntry);
            XDocument relationshipsDocument = ReadXml(relationshipsEntry);
            bool slideChanged = false;
            var foundShapeIds = new HashSet<int>();

            foreach (XElement picture in slideDocument.Descendants(PresentationNamespace + "pic"))
            {
                XElement? nonVisualProperties = picture.Element(PresentationNamespace + "nvPicPr")
                    ?.Element(PresentationNamespace + "cNvPr");
                if (!int.TryParse((string?)nonVisualProperties?.Attribute("id"), out int shapeId) ||
                    !shapeIds.Contains(shapeId))
                {
                    continue;
                }

                foundShapeIds.Add(shapeId);
                cancellationToken.ThrowIfCancellationRequested();

                XElement? blip = picture.Element(PresentationNamespace + "blipFill")
                    ?.Element(DrawingNamespace + "blip");
                string shapeLabel = $"slide {slideIndex}, picture id {shapeId}";
                if (picture.Descendants().Any(element => element.Name.LocalName == "svgBlip"))
                {
                    skipped.Add($"{shapeLabel}: pictures with an SVG version are left unchanged.");
                    continue;
                }

                if (blip is null)
                {
                    skipped.Add($"{shapeLabel}: picture has no image data.");
                    continue;
                }

                string? relationshipId = (string?)blip.Attribute(OfficeRelationshipNamespace + "embed");
                if (relationshipId is null || blip.Attribute(OfficeRelationshipNamespace + "link") is not null)
                {
                    skipped.Add($"{shapeLabel}: external linked pictures are not changed.");
                    continue;
                }

                XElement? relationship = relationshipsDocument
                    .Root?
                    .Elements(PackageRelationshipNamespace + "Relationship")
                    .SingleOrDefault(element => (string?)element.Attribute("Id") == relationshipId);
                string? relationshipTarget = (string?)relationship?.Attribute("Target");
                if (relationship is null || string.IsNullOrWhiteSpace(relationshipTarget))
                {
                    throw new InvalidDataException(
                        $"Picture {shapeLabel} refers to missing package relationship '{relationshipId}'.");
                }

                string mediaPath = ResolvePartPath(slidePath, relationshipTarget);
                ZipArchiveEntry mediaEntry = RequireEntry(source, mediaPath);
                string extension = Path.GetExtension(mediaPath).ToLowerInvariant();
                if (!IsSupportedRasterExtension(extension))
                {
                    skipped.Add($"{shapeLabel}: '{extension}' is not a supported raster format and was left unchanged.");
                    continue;
                }

                byte[] originalImage = ReadBytes(mediaEntry);
                PictureTransform transform;
                try
                {
                    transform = TransformImage(
                        originalImage,
                        extension,
                        GetCrop(picture),
                        GetPictureSize(picture),
                        targetPpi,
                        deleteCroppedAreas);
                }
                catch (ArgumentException)
                {
                    skipped.Add($"{shapeLabel}: the image format could not be decoded.");
                    continue;
                }
                catch (ExternalException)
                {
                    skipped.Add($"{shapeLabel}: the image could not be decoded by Windows imaging.");
                    continue;
                }

                if (!transform.Changed)
                {
                    if (transform.SkipReason is not null)
                    {
                        skipped.Add($"{shapeLabel}: {transform.SkipReason}");
                    }
                    continue;
                }

                string newMediaName = $"image-mcp-{Guid.NewGuid():N}{extension}";
                string newMediaPath = $"ppt/media/{newMediaName}";
                addedParts.Add(newMediaPath, transform.Bytes);

                string newRelationshipId = NextRelationshipId(relationshipsDocument);
                relationship.AddBeforeSelf(new XElement(
                    PackageRelationshipNamespace + "Relationship",
                    new XAttribute("Id", newRelationshipId),
                    new XAttribute(
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                    new XAttribute("Target", $"../media/{newMediaName}")));
                blip.SetAttributeValue(OfficeRelationshipNamespace + "embed", newRelationshipId);
                blip.Attribute(OfficeRelationshipNamespace + "link")?.Remove();
                if (!slideDocument.Descendants(DrawingNamespace + "blip")
                    .Any(otherBlip => (string?)otherBlip.Attribute(OfficeRelationshipNamespace + "embed") == relationshipId))
                {
                    relationship.Remove();
                }

                if (deleteCroppedAreas)
                {
                    picture.Element(PresentationNamespace + "blipFill")
                        ?.Element(DrawingNamespace + "srcRect")
                        ?.Remove();
                }

                slideChanged = true;
                changedImageParts.Add(mediaPath);
                compressedCount++;
                originalBytes += originalImage.LongLength;
                compressedBytes += transform.Bytes.LongLength;
            }

            foreach (int missingShapeId in shapeIds.Except(foundShapeIds))
            {
                throw new InvalidDataException(
                    $"Picture id {missingShapeId} on slide {slideIndex} was not found in the saved presentation package.");
            }

            if (slideChanged)
            {
                replacements[slidePath] = Serialize(slideDocument);
                replacements[relationshipsPath] = Serialize(relationshipsDocument);
            }
        }

        if (compressedCount == 0)
        {
            return new PresentationCompressionResult(
                resolution, compressedCount, skipped, originalBytes, compressedBytes);
        }

        var unreferencedMedia = changedImageParts
            .Where(mediaPath => !IsReferenced(source, replacements, mediaPath))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        string outputPath = Path.Combine(
            Path.GetDirectoryName(presentationPath)!,
            $".{Path.GetFileNameWithoutExtension(presentationPath)}-{Guid.NewGuid():N}.zip");

        try
        {
            using (var output = ZipFile.Open(outputPath, ZipArchiveMode.Create))
            {
                foreach (ZipArchiveEntry entry in source.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (unreferencedMedia.Contains(entry.FullName))
                    {
                        continue;
                    }

                    ZipArchiveEntry newEntry = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    newEntry.LastWriteTime = entry.LastWriteTime;
                    using Stream destination = newEntry.Open();
                    if (replacements.TryGetValue(entry.FullName, out byte[]? replacement))
                    {
                        destination.Write(replacement);
                    }
                    else
                    {
                        using Stream original = entry.Open();
                        original.CopyTo(destination);
                    }
                }

                foreach ((string partPath, byte[] bytes) in addedParts)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ZipArchiveEntry newEntry = output.CreateEntry(partPath, CompressionLevel.Optimal);
                    using Stream destination = newEntry.Open();
                    destination.Write(bytes);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            File.Replace(outputPath, presentationPath, destinationBackupFileName: null);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }

        return new PresentationCompressionResult(
            resolution, compressedCount, skipped, originalBytes, compressedBytes);
    }

    private static PictureTransform TransformImage(
        byte[] originalBytes,
        string extension,
        CropMargins crop,
        (long Width, long Height)? frameSize,
        int? targetPpi,
        bool deleteCroppedAreas)
    {
        if (deleteCroppedAreas &&
            (crop.Left < 0d || crop.Top < 0d || crop.Right < 0d || crop.Bottom < 0d))
        {
            return PictureTransform.Skip(
                "negative crop margins expand beyond the source image and cannot be removed without changing the visible framing.");
        }

        using var input = new MemoryStream(originalBytes, writable: false);
        using System.Drawing.Image source =
            System.Drawing.Image.FromStream(input, useEmbeddedColorManagement: true, validateImageData: true);
        if (source.FrameDimensionsList.Any(dimension => source.GetFrameCount(new FrameDimension(dimension)) > 1))
        {
            return PictureTransform.Skip("animated or multi-frame images are not modified.");
        }

        int left = deleteCroppedAreas ? PixelMargin(source.Width, crop.Left) : 0;
        int top = deleteCroppedAreas ? PixelMargin(source.Height, crop.Top) : 0;
        int right = deleteCroppedAreas ? PixelMargin(source.Width, crop.Right) : 0;
        int bottom = deleteCroppedAreas ? PixelMargin(source.Height, crop.Bottom) : 0;
        int visibleWidth = source.Width - left - right;
        int visibleHeight = source.Height - top - bottom;
        if (visibleWidth <= 0 || visibleHeight <= 0)
        {
            return PictureTransform.Skip("crop margins leave no visible image area.");
        }

        int targetWidth = visibleWidth;
        int targetHeight = visibleHeight;
        if (targetPpi.HasValue)
        {
            if (frameSize is null || frameSize.Value.Width <= 0 || frameSize.Value.Height <= 0)
            {
                return PictureTransform.Skip("PowerPoint did not expose a usable crop-frame size.");
            }

            double cropWidth = deleteCroppedAreas ? 1d : 1d - crop.Left - crop.Right;
            double cropHeight = deleteCroppedAreas ? 1d : 1d - crop.Top - crop.Bottom;
            if (cropWidth <= 0d || cropHeight <= 0d)
            {
                return PictureTransform.Skip("crop margins leave no visible image area.");
            }

            double widthMultiplier = deleteCroppedAreas ? 1d : 1d / cropWidth;
            double heightMultiplier = deleteCroppedAreas ? 1d : 1d / cropHeight;
            int requestedWidth = (int)Math.Max(1, Math.Round(
                frameSize.Value.Width / EmusPerInch * targetPpi.Value * widthMultiplier));
            int requestedHeight = (int)Math.Max(1, Math.Round(
                frameSize.Value.Height / EmusPerInch * targetPpi.Value * heightMultiplier));
            double scale = Math.Min(
                1d,
                Math.Min((double)requestedWidth / visibleWidth, (double)requestedHeight / visibleHeight));
            targetWidth = Math.Max(1, (int)Math.Round(visibleWidth * scale));
            targetHeight = Math.Max(1, (int)Math.Round(visibleHeight * scale));
        }

        bool cropApplied = deleteCroppedAreas && (left != 0 || top != 0 || right != 0 || bottom != 0);
        bool dimensionsChanged = targetWidth != visibleWidth || targetHeight != visibleHeight;
        if (!cropApplied && !dimensionsChanged)
        {
            return PictureTransform.NoChange;
        }

        using var bitmap = new Bitmap(targetWidth, targetHeight, PixelFormat.Format32bppArgb);
        if (targetPpi.HasValue)
        {
            bitmap.SetResolution(targetPpi.Value, targetPpi.Value);
        }

        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.Transparent);
            graphics.CompositingMode = CompositingMode.SourceCopy;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(
                source,
                new Rectangle(0, 0, targetWidth, targetHeight),
                new Rectangle(left, top, visibleWidth, visibleHeight),
                GraphicsUnit.Pixel);
        }

        byte[] encoded = Encode(bitmap, extension, targetPpi.HasValue ? 88L : 95L);
        if (encoded.Length >= originalBytes.Length)
        {
            return PictureTransform.Skip("re-encoding would not reduce the image data size.");
        }

        return new PictureTransform(encoded, true, null);
    }

    private static byte[] Encode(Bitmap bitmap, string extension, long jpegQuality)
    {
        using var output = new MemoryStream();
        if (extension is ".jpg" or ".jpeg")
        {
            ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders()
                .Single(encoder => encoder.FormatID == ImageFormat.Jpeg.Guid);
            using var parameters = new EncoderParameters(1);
            parameters.Param[0] = new EncoderParameter(Encoder.Quality, jpegQuality);
            bitmap.Save(output, codec, parameters);
        }
        else
        {
            ImageFormat format = extension switch
            {
                ".png" => ImageFormat.Png,
                ".bmp" => ImageFormat.Bmp,
                ".gif" => ImageFormat.Gif,
                ".tif" or ".tiff" => ImageFormat.Tiff,
                _ => throw new ArgumentOutOfRangeException(nameof(extension), extension, "Unsupported raster format.")
            };
            bitmap.Save(output, format);
        }

        return output.ToArray();
    }

    private static bool IsSupportedRasterExtension(string extension) =>
        extension is ".png" or ".jpg" or ".jpeg" or ".bmp" or ".gif" or ".tif" or ".tiff";

    private static int PixelMargin(int dimension, double fraction) =>
        (int)Math.Clamp(Math.Round(dimension * fraction), 0, dimension - 1);

    private static CropMargins GetCrop(XElement picture)
    {
        XElement? crop = picture.Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "srcRect");
        return new CropMargins(
            ReadCropPercent(crop, "l"),
            ReadCropPercent(crop, "t"),
            ReadCropPercent(crop, "r"),
            ReadCropPercent(crop, "b"));
    }

    private static double ReadCropPercent(XElement? crop, string side)
    {
        if (!int.TryParse((string?)crop?.Attribute(side), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
        {
            return 0d;
        }
        return value / 100000d;
    }

    private static (long Width, long Height)? GetPictureSize(XElement picture)
    {
        XElement? extent = picture.Element(PresentationNamespace + "spPr")
            ?.Element(DrawingNamespace + "xfrm")
            ?.Element(DrawingNamespace + "ext");
        if (!long.TryParse((string?)extent?.Attribute("cx"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long width) ||
            !long.TryParse((string?)extent?.Attribute("cy"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long height))
        {
            return null;
        }
        return (width, height);
    }

    private static List<string> GetSlidePaths(ZipArchive archive)
    {
        XDocument presentation = ReadXml(RequireEntry(archive, "ppt/presentation.xml"));
        XDocument relationships = ReadXml(RequireEntry(archive, "ppt/_rels/presentation.xml.rels"));
        var relationshipTargets = relationships.Root?
            .Elements(PackageRelationshipNamespace + "Relationship")
            .ToDictionary(
                element => (string?)element.Attribute("Id") ?? "",
                element => (string?)element.Attribute("Target") ?? "",
                StringComparer.Ordinal);
        if (relationshipTargets is null)
        {
            throw new InvalidDataException("Presentation package has no relationship table.");
        }

        var result = new List<string>();
        foreach (XElement slideId in presentation.Descendants(PresentationNamespace + "sldId"))
        {
            string relationshipId = (string?)slideId.Attribute(OfficeRelationshipNamespace + "id") ?? "";
            if (!relationshipTargets.TryGetValue(relationshipId, out string? target) ||
                string.IsNullOrWhiteSpace(target))
            {
                throw new InvalidDataException(
                    $"Presentation slide reference '{relationshipId}' has no target.");
            }
            result.Add(ResolvePartPath("ppt/presentation.xml", target));
        }

        return result;
    }

    private static bool IsReferenced(
        ZipArchive archive,
        Dictionary<string, byte[]> replacements,
        string mediaPath)
    {
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry =>
                     entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            XDocument relationships = replacements.TryGetValue(entry.FullName, out byte[]? replacement)
                ? ReadXml(replacement)
                : ReadXml(entry);
            string ownerPath = GetRelationshipOwnerPath(entry.FullName);
            foreach (XElement relationship in relationships.Descendants(PackageRelationshipNamespace + "Relationship"))
            {
                if (!string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        ResolvePartPath(ownerPath, (string?)relationship.Attribute("Target") ?? ""),
                        mediaPath,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        return false;
    }

    private static string GetRelationshipOwnerPath(string relationshipsPath)
    {
        string directory = Path.GetDirectoryName(relationshipsPath.Replace('/', '\\')) ?? "";
        string ownerDirectory = directory.EndsWith("\\_rels", StringComparison.OrdinalIgnoreCase)
            ? directory[..^6]
            : directory;
        string relationshipFile = Path.GetFileName(relationshipsPath);
        string ownerFile = relationshipFile.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
            ? relationshipFile[..^5]
            : relationshipFile;
        return $"{ownerDirectory.Replace('\\', '/')}/{ownerFile}".TrimStart('/');
    }

    private static string GetRelationshipsPath(string partPath)
    {
        int lastSlash = partPath.LastIndexOf('/');
        return $"{partPath[..(lastSlash + 1)]}_rels/{partPath[(lastSlash + 1)..]}.rels";
    }

    private static string ResolvePartPath(string ownerPartPath, string target)
    {
        if (target.StartsWith('/'))
        {
            return NormalizePartPath(target.TrimStart('/'));
        }

        string directory = ownerPartPath.Contains('/')
            ? ownerPartPath[..(ownerPartPath.LastIndexOf('/') + 1)]
            : "";
        return NormalizePartPath(directory + target);
    }

    private static string NormalizePartPath(string path)
    {
        var parts = new List<string>();
        foreach (string part in path.Replace('\\', '/').Split('/'))
        {
            if (part is "" or ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0)
                {
                    throw new InvalidDataException("Package relationship escapes the Open XML package root.");
                }
                parts.RemoveAt(parts.Count - 1);
            }
            else
            {
                parts.Add(part);
            }
        }
        return string.Join('/', parts);
    }

    private static string NextRelationshipId(XDocument relationships)
    {
        var usedIds = relationships.Root?
            .Elements(PackageRelationshipNamespace + "Relationship")
            .Select(element => (string?)element.Attribute("Id"))
            .ToHashSet(StringComparer.Ordinal) ?? [];
        int suffix = 1;
        string candidate;
        do
        {
            candidate = $"rIdMcpImage{suffix++}";
        } while (usedIds.Contains(candidate));
        return candidate;
    }

    private static ZipArchiveEntry RequireEntry(ZipArchive archive, string path) =>
        archive.GetEntry(path) ?? throw new InvalidDataException(
            $"Presentation package is missing required part '{path}'.");

    private static XDocument ReadXml(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static XDocument ReadXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadBytes(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        using var output = new MemoryStream();
        stream.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] Serialize(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = System.Xml.XmlWriter.Create(stream, new System.Xml.XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            CloseOutput = false
        }))
        {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    private readonly record struct CropMargins(
        double Left,
        double Top,
        double Right,
        double Bottom);

    private sealed record PictureTransform(byte[] Bytes, bool Changed, string? SkipReason)
    {
        public static PictureTransform NoChange { get; } = new([], false, null);
        public static PictureTransform Skip(string reason) => new([], false, reason);
    }
}
