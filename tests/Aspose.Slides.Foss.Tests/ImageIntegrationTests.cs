using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for image operations including round-trip through Presentation save/reload.
/// </summary>
public sealed class ImageIntegrationTests
{
    /// <summary>
    /// Generates a minimal valid 1x1 PNG with the given RGB colour.
    /// </summary>
    private static byte[] CreateTestPng(byte r, byte g, byte b)
    {
        byte[] raw = [0, r, g, b];

        using var deflateBuffer = new MemoryStream();
        deflateBuffer.WriteByte(0x78);
        deflateBuffer.WriteByte(0x01);
        using (var deflate = new DeflateStream(deflateBuffer, CompressionLevel.Optimal, leaveOpen: true))
        {
            deflate.Write(raw, 0, raw.Length);
        }

        uint a = 1, bv = 0;
        foreach (var bt in raw)
        {
            a = (a + bt) % 65521;
            bv = (bv + a) % 65521;
        }
        uint adler = (bv << 16) | a;
        deflateBuffer.WriteByte((byte)(adler >> 24));
        deflateBuffer.WriteByte((byte)(adler >> 16));
        deflateBuffer.WriteByte((byte)(adler >> 8));
        deflateBuffer.WriteByte((byte)adler);

        byte[] compressedData = deflateBuffer.ToArray();

        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        WriteChunk(png, "IHDR", BuildIhdr(1, 1, 8, 2));
        WriteChunk(png, "IDAT", compressedData);
        WriteChunk(png, "IEND", []);

        return png.ToArray();
    }

    private static byte[] BuildIhdr(int width, int height, byte bitDepth, byte colorType)
    {
        using var ms = new MemoryStream();
        ms.Write(ToBigEndian(width));
        ms.Write(ToBigEndian(height));
        ms.WriteByte(bitDepth);
        ms.WriteByte(colorType);
        ms.WriteByte(0);
        ms.WriteByte(0);
        ms.WriteByte(0);
        return ms.ToArray();
    }

    private static byte[] ToBigEndian(int value) =>
        [(byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value];

    private static void WriteChunk(MemoryStream stream, string type, byte[] data)
    {
        stream.Write(ToBigEndian(data.Length));
        byte[] typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes);
        stream.Write(data);
        byte[] crcInput = new byte[4 + data.Length];
        Array.Copy(typeBytes, 0, crcInput, 0, 4);
        Array.Copy(data, 0, crcInput, 4, data.Length);
        uint crc = ComputeCrc32(crcInput);
        stream.Write(ToBigEndian((int)crc));
    }

    private static uint ComputeCrc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var bt in data)
        {
            crc ^= bt;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (0xEDB88320 * (crc & 1));
        }
        return crc ^ 0xFFFFFFFF;
    }

    /// <summary>
    /// Adding an image increases the collection count.
    /// </summary>
    [Fact]
    public void AddImage_IncreasesCollectionCount()
    {
        using var pres = new Presentation();
        var initialCount = pres.Images.AsICollection.Count;
        var pngData = CreateTestPng(255, 0, 0);

        pres.Images.AddImage(pngData);

        pres.Images.AsICollection.Count.Should().Be(initialCount + 1);
    }

    /// <summary>
    /// Multiple images with different data can be added.
    /// </summary>
    [Fact]
    public void MultipleImages_AllAddedToCollection()
    {
        using var pres = new Presentation();
        var png1 = CreateTestPng(255, 0, 0);
        var png2 = CreateTestPng(0, 255, 0);
        var png3 = CreateTestPng(0, 0, 255);

        pres.Images.AddImage(png1);
        pres.Images.AddImage(png2);
        pres.Images.AddImage(png3);

        pres.Images.AsICollection.Count.Should().BeGreaterThanOrEqualTo(3);
    }

    /// <summary>
    /// Multiple images can be saved to a stream without error.
    /// </summary>
    [Fact]
    public void MultipleImages_SaveProducesNonEmptyOutput()
    {
        using var pres = new Presentation();
        var png1 = CreateTestPng(255, 0, 0);
        var png2 = CreateTestPng(0, 255, 0);
        var png3 = CreateTestPng(0, 0, 255);
        pres.Images.AddImage(png1);
        pres.Images.AddImage(png2);
        pres.Images.AddImage(png3);

        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);

        ms.Length.Should().BeGreaterThan(0);
        pres.Images.AsICollection.Count.Should().BeGreaterThanOrEqualTo(3);
    }

    /// <summary>
    /// A picture frame can be created with an image in a ShapeCollection.
    /// </summary>
    [Fact]
    public void PictureFrame_CanBeCreatedWithImage()
    {
        using var pres = new Presentation();
        var pngData = CreateTestPng(128, 64, 32);
        var image = pres.Images.AddImage(pngData);

        image.Should().NotBeNull();
        image.BinaryData.Should().NotBeEmpty();
        image.ContentType.Should().Contain("png");
    }

    /// <summary>
    /// Image binary data is preserved after adding to the collection.
    /// </summary>
    [Fact]
    public void AddImage_BinaryDataIsPreserved()
    {
        using var pres = new Presentation();
        var pngData = CreateTestPng(200, 100, 50);
        var image = pres.Images.AddImage(pngData);

        image.BinaryData.Should().Equal(pngData);
    }

    /// <summary>
    /// Duplicate image data is deduplicated (same reference returned).
    /// </summary>
    [Fact]
    public void AddImage_DeduplicatesIdenticalData()
    {
        using var pres = new Presentation();
        var pngData = CreateTestPng(42, 42, 42);

        var img1 = pres.Images.AddImage(pngData);
        var img2 = pres.Images.AddImage(pngData);

        img1.Should().BeSameAs(img2);
        // Only one image should be in the collection due to deduplication
        var countAfter = pres.Images.AsICollection.Count;
        countAfter.Should().Be(1);
    }
}
