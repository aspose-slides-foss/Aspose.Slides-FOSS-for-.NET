using System.IO.Compression;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests ImageCollection, PPImage, Image, and PictureFrame operations.
/// </summary>
public sealed class ImageTests
{
    private static uint ComputeCrc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (0xEDB88320 * (crc & 1));
        }
        return crc ^ 0xFFFFFFFF;
    }

    /// <summary>
    /// Generates a minimal valid 1x1 PNG with the given RGB colour.
    /// </summary>
    private static byte[] CreateTestPng(byte r = 255, byte g = 0, byte b = 0)
    {
        static byte[] ToBigEndian(int value)
        {
            var bytes = BitConverter.GetBytes(value);
            if (BitConverter.IsLittleEndian) Array.Reverse(bytes);
            return bytes;
        }

        byte[] Chunk(byte[] chunkType, byte[] data)
        {
            var combined = new byte[chunkType.Length + data.Length];
            Buffer.BlockCopy(chunkType, 0, combined, 0, chunkType.Length);
            Buffer.BlockCopy(data, 0, combined, chunkType.Length, data.Length);

            var crc = ComputeCrc32(combined);

            using var ms = new MemoryStream();
            ms.Write(ToBigEndian(data.Length));
            ms.Write(combined);
            ms.Write(ToBigEndian(unchecked((int)crc)));
            return ms.ToArray();
        }

        // IHDR: 1x1, 8-bit depth, RGB color type (2)
        using var ihdrData = new MemoryStream();
        ihdrData.Write(ToBigEndian(1)); // width
        ihdrData.Write(ToBigEndian(1)); // height
        ihdrData.WriteByte(8);          // bit depth
        ihdrData.WriteByte(2);          // color type (RGB)
        ihdrData.WriteByte(0);          // compression
        ihdrData.WriteByte(0);          // filter
        ihdrData.WriteByte(0);          // interlace

        // Raw scanline: filter byte + RGB
        byte[] raw = [0, r, g, b];
        byte[] compressed;
        using (var compMs = new MemoryStream())
        {
            using (var deflate = new ZLibStream(compMs, CompressionLevel.Optimal, leaveOpen: true))
            {
                deflate.Write(raw);
            }
            compressed = compMs.ToArray();
        }

        byte[] header = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        using var result = new MemoryStream();
        result.Write(header);
        result.Write(Chunk("IHDR"u8.ToArray(), ihdrData.ToArray()));
        result.Write(Chunk("IDAT"u8.ToArray(), compressed));
        result.Write(Chunk("IEND"u8.ToArray(), Array.Empty<byte>()));
        return result.ToArray();
    }

    /// <summary>
    /// Creates an initialized ImageCollection backed by an in-memory OPC package.
    /// </summary>
    private static ImageCollection CreateImageCollection()
    {
        var package = new OpcPackage();
        var collection = new ImageCollection();
        collection.InitInternal(package);
        return collection;
    }

    /// <summary>
    /// Adding an image increases collection count.
    /// </summary>
    [Fact]
    public void AddImage_IncreasesCount()
    {
        var collection = CreateImageCollection();
        var pngBytes = CreateTestPng(255, 0, 0);

        collection.AddImage(pngBytes);

        collection.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// Added image can be retrieved by index.
    /// </summary>
    [Fact]
    public void AddImage_RetrievableByIndex()
    {
        var collection = CreateImageCollection();
        var pngBytes = CreateTestPng(255, 0, 0);

        var ppImage = collection.AddImage(pngBytes);

        collection[0].Should().BeSameAs(ppImage);
    }

    /// <summary>
    /// Multiple images can be added and iterated.
    /// </summary>
    [Fact]
    public void MultipleImages_CanBeAddedAndIterated()
    {
        var collection = CreateImageCollection();
        collection.AddImage(CreateTestPng(255, 0, 0));
        collection.AddImage(CreateTestPng(0, 255, 0));
        collection.AddImage(CreateTestPng(0, 0, 255));

        collection.Count.Should().BeGreaterThanOrEqualTo(3);
        var images = collection.ToList();
        images.Should().HaveCountGreaterThanOrEqualTo(3);
    }

    /// <summary>
    /// Each added image is a distinct object.
    /// </summary>
    [Fact]
    public void MultipleImages_AreDistinctObjects()
    {
        var collection = CreateImageCollection();
        var img1 = collection.AddImage(CreateTestPng(255, 0, 0));
        var img2 = collection.AddImage(CreateTestPng(0, 255, 0));
        var img3 = collection.AddImage(CreateTestPng(0, 0, 255));

        img1.Should().NotBeSameAs(img2);
        img2.Should().NotBeSameAs(img3);
    }

    /// <summary>
    /// Added image has non-empty binary data.
    /// </summary>
    [Fact]
    public void AddImage_BinaryDataIsNotEmpty()
    {
        var collection = CreateImageCollection();
        var pngBytes = CreateTestPng(255, 0, 0);

        var ppImage = collection.AddImage(pngBytes);

        ppImage.BinaryData.Should().NotBeEmpty();
    }

    /// <summary>
    /// Added image has PNG content type.
    /// </summary>
    [Fact]
    public void AddImage_HasPngContentType()
    {
        var collection = CreateImageCollection();
        var pngBytes = CreateTestPng(255, 0, 0);

        var ppImage = collection.AddImage(pngBytes);

        ppImage.ContentType.Should().Contain("png");
    }

    /// <summary>
    /// PictureFrame can be instantiated.
    /// </summary>
    [Fact]
    public void PictureFrame_CanBeInstantiated()
    {
        var pf = new PictureFrame();
        pf.Should().NotBeNull();
    }

    /// <summary>
    /// ShapeType.Rectangle is defined (used for picture frames).
    /// </summary>
    [Fact]
    public void PictureFrame_RectangleShapeTypeIsDefined()
    {
        Enum.IsDefined(ShapeType.Rectangle).Should().BeTrue();
        ShapeType.Rectangle.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// Images.FromStream creates a valid Image from PNG data.
    /// </summary>
    [Fact]
    public void ImageFromStream_CreatesValidImage()
    {
        var pngBytes = CreateTestPng(0, 128, 255);
        using var stream = new MemoryStream(pngBytes);

        var image = Images.FromStream(stream);

        image.Should().NotBeNull();
        image.Width.Should().Be(1);
        image.Height.Should().Be(1);
    }

    /// <summary>
    /// Images.FromFile creates a valid Image from a temporary PNG file.
    /// </summary>
    [Fact]
    public void ImageFromFile_CreatesValidImage()
    {
        var pngBytes = CreateTestPng(0, 0, 255);
        var tempFile = Path.GetTempFileName() + ".png";
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);

            var image = Images.FromFile(tempFile);

            image.Should().NotBeNull();
            image.Width.Should().Be(1);
            image.Height.Should().Be(1);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    /// <summary>
    /// Image loaded from file can be added to ImageCollection.
    /// </summary>
    [Fact]
    public void ImageFromFile_CanBeAddedToCollection()
    {
        var pngBytes = CreateTestPng(0, 0, 255);
        var tempFile = Path.GetTempFileName() + ".png";
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var image = Images.FromFile(tempFile);
            var collection = CreateImageCollection();

            var ppImage = collection.AddImage(image);

            ppImage.Should().NotBeNull();
            collection.Count.Should().BeGreaterThanOrEqualTo(1);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    /// <summary>
    /// AddImage via stream overload works correctly.
    /// </summary>
    [Fact]
    public void AddImage_ViaStream_IncreasesCount()
    {
        var collection = CreateImageCollection();
        var pngBytes = CreateTestPng(128, 0, 128);
        using var stream = new MemoryStream(pngBytes);

        collection.AddImage(stream);

        collection.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// PPImage exposes Width and Height.
    /// </summary>
    [Fact]
    public void PPImage_ExposesWidthAndHeight()
    {
        var collection = CreateImageCollection();
        var ppImage = collection.AddImage(CreateTestPng(255, 0, 0));

        ppImage.Width.Should().Be(1);
        ppImage.Height.Should().Be(1);
    }

    /// <summary>
    /// PPImage.Image returns an IImage wrapper.
    /// </summary>
    [Fact]
    public void PPImage_ImagePropertyReturnsWrapper()
    {
        var collection = CreateImageCollection();
        var ppImage = collection.AddImage(CreateTestPng(0, 255, 0));

        var image = ppImage.Image;

        image.Should().NotBeNull();
        image.Width.Should().Be(1);
        image.Height.Should().Be(1);
    }

    /// <summary>
    /// Image.Save writes data to a stream.
    /// </summary>
    [Fact]
    public void Image_SaveToStream_WritesData()
    {
        var pngBytes = CreateTestPng(255, 128, 0);
        using var input = new MemoryStream(pngBytes);
        var image = Images.FromStream(input);

        using var output = new MemoryStream();
        image.Save(output, "png");

        output.Length.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// AsICollection returns all images.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsAllImages()
    {
        var collection = CreateImageCollection();
        collection.AddImage(CreateTestPng(255, 0, 0));
        collection.AddImage(CreateTestPng(0, 255, 0));

        var list = collection.AsICollection;

        list.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// AsIEnumerable returns all images.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsAllImages()
    {
        var collection = CreateImageCollection();
        collection.AddImage(CreateTestPng(255, 0, 0));
        collection.AddImage(CreateTestPng(0, 255, 0));

        var enumerable = collection.AsIEnumerable.ToList();

        enumerable.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// Image.Size returns correct dimensions.
    /// </summary>
    [Fact]
    public void Image_SizeProperty_ReturnsCorrectDimensions()
    {
        var pngBytes = CreateTestPng(0, 0, 0);
        using var stream = new MemoryStream(pngBytes);
        var image = Images.FromStream(stream);

        image.Size.Width.Should().Be(1);
        image.Size.Height.Should().Be(1);
    }

    /// <summary>
    /// PPImage.ReplaceImage updates the binary data.
    /// </summary>
    [Fact]
    public void PPImage_ReplaceImage_UpdatesData()
    {
        var collection = CreateImageCollection();
        var ppImage = collection.AddImage(CreateTestPng(255, 0, 0));
        var originalData = ppImage.BinaryData;

        var newPng = CreateTestPng(0, 255, 0);
        ppImage.ReplaceImage(newPng);

        ppImage.BinaryData.Should().NotBeEquivalentTo(originalData);
    }
}
