using System.Collections;
using System.Text.RegularExpressions;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Concrete collection managing presentation images within an OPC package.
/// </summary>
public sealed partial class ImageCollection : IImageCollection, IEnumerable<IPPImage>
{
    private OpcPackage? _package;
    private List<PPImage>? _images;

    /// <summary>
    /// Initializes internal state by scanning the package for existing media parts.
    /// </summary>
    internal void InitInternal(OpcPackage package)
    {
        _package = package;
        _images = new List<PPImage>();

        var partNames = package.GetSortedPartNames()
            .Where(n => n.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var partName in partNames)
        {
            var data = package.GetPart(partName);
            if (data is null)
                continue;

            var contentType = ImageHelpers.GuessContentType(data);
            var ppImage = new PPImage();
            ppImage.InitInternal(package, partName, data, contentType);
            _images.Add(ppImage);
        }
    }

    /// <inheritdoc/>
    public override IList<IPPImage> AsICollection =>
        new List<IPPImage>(_images ?? Enumerable.Empty<PPImage>());

    /// <inheritdoc/>
    public override IEnumerable<IPPImage> AsIEnumerable =>
        _images ?? Enumerable.Empty<PPImage>();

    /// <inheritdoc/>
    public override IPPImage AddImage(IImage image)
    {
        if (image is Image img)
            return AddImageCore(img.Data);

        throw new ArgumentException("Unsupported IImage implementation.", nameof(image));
    }

    /// <inheritdoc/>
    public override IPPImage AddImage(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return AddImageCore(ms.ToArray());
    }

    /// <inheritdoc/>
    public override IPPImage AddImage(byte[] buffer)
    {
        return AddImageCore(buffer);
    }

    /// <inheritdoc/>
    public override IPPImage this[int index] =>
        (_images ?? throw new InvalidOperationException("Collection not initialized."))[index];

    /// <summary>
    /// Gets the number of images in the collection.
    /// </summary>
    public override int Count => _images?.Count ?? 0;

    /// <summary>
    /// Returns an enumerator that iterates through the images.
    /// </summary>
    public IEnumerator<IPPImage> GetEnumerator()
    {
        if (_images is null)
            yield break;

        foreach (var img in _images)
            yield return img;
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Finds an existing <see cref="PPImage"/> by its part name.
    /// </summary>
    internal PPImage? FindByPartName(string partName)
    {
        return _images?.Find(img =>
            string.Equals(img.PartName, partName, StringComparison.OrdinalIgnoreCase));
    }

    private IPPImage AddImageCore(byte[] data)
    {
        if (_package is null || _images is null)
            throw new InvalidOperationException("Collection not initialized.");

        // Deduplication: return existing if identical data found.
        foreach (var existing in _images)
        {
            if (existing.ImageData.AsSpan().SequenceEqual(data))
                return existing;
        }

        var contentType = ImageHelpers.GuessContentType(data);
        var extension = ImageHelpers.GetExtensionForContentType(contentType);
        var nextNumber = FindNextImageNumber();
        var partName = $"ppt/media/image{nextNumber}.{extension}";

        _package.SetPart(partName, data);

        // Without a Default for this extension the media part resolves no content type at all, and
        // per ISO/IEC 29500-2 10.1.2 the content type is the part's identity: PowerPoint rejects the
        // package. This has to be written even when nothing ends up referencing the image.
        OpcRegistration.AddContentTypeDefault(_package, extension, contentType);

        var ppImage = new PPImage();
        ppImage.InitInternal(_package, partName, data, contentType);
        _images.Add(ppImage);
        return ppImage;
    }

    private int FindNextImageNumber()
    {
        if (_package is null)
            return 1;

        var usedNumbers = new HashSet<int>();
        var regex = ImageNumberRegex();

        foreach (var name in _package.GetPartNames())
        {
            var match = regex.Match(name);
            if (match.Success && int.TryParse(match.Groups[1].Value, out var num))
            {
                usedNumbers.Add(num);
            }
        }

        for (int i = 1; ; i++)
        {
            if (!usedNumbers.Contains(i))
                return i;
        }
    }

    [GeneratedRegex(@"ppt/media/image(\d+)\.", RegexOptions.IgnoreCase)]
    private static partial Regex ImageNumberRegex();
}
