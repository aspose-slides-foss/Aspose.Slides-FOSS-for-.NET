using System.Buffers.Binary;
using System.Collections.Frozen;

namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// Image utility functions for parsing image headers and detecting content types.
/// Supports JPEG, PNG, GIF, BMP, TIFF, EMF, and WMF formats without external dependencies.
/// </summary>
internal static class ImageUtils
{
    // Magic byte signatures for image format detection
    private static ReadOnlySpan<byte> PngSignature => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    private static ReadOnlySpan<byte> JpegSignature => [0xFF, 0xD8, 0xFF];
    private static ReadOnlySpan<byte> Gif87Signature => "GIF87a"u8;
    private static ReadOnlySpan<byte> Gif89Signature => "GIF89a"u8;
    private static ReadOnlySpan<byte> BmpSignature => "BM"u8;
    private static ReadOnlySpan<byte> TiffLeSignature => [0x49, 0x49, 0x2A, 0x00];
    private static ReadOnlySpan<byte> TiffBeSignature => [0x4D, 0x4D, 0x00, 0x2A];
    private static ReadOnlySpan<byte> EmfSignature => [0x01, 0x00, 0x00, 0x00];
    private static ReadOnlySpan<byte> WmfSignature => [0xD7, 0xCD, 0xC6, 0x9A];

    /// <summary>
    /// Extension to MIME type mapping (without leading dot).
    /// </summary>
    public static FrozenDictionary<string, string> ExtensionContentTypes { get; } = new Dictionary<string, string>
    {
        ["png"] = "image/png",
        ["jpg"] = "image/jpeg",
        ["jpeg"] = "image/jpeg",
        ["gif"] = "image/gif",
        ["bmp"] = "image/bmp",
        ["tiff"] = "image/tiff",
        ["tif"] = "image/tiff",
        ["emf"] = "image/x-emf",
        ["wmf"] = "image/x-wmf",
        ["svg"] = "image/svg+xml",
    }.ToFrozenDictionary();

    /// <summary>
    /// MIME type to default extension mapping (without leading dot).
    /// </summary>
    public static FrozenDictionary<string, string> ContentTypeExtensions { get; } = new Dictionary<string, string>
    {
        ["image/png"] = "png",
        ["image/jpeg"] = "jpeg",
        ["image/gif"] = "gif",
        ["image/bmp"] = "bmp",
        ["image/tiff"] = "tiff",
        ["image/x-emf"] = "emf",
        ["image/x-wmf"] = "wmf",
        ["image/svg+xml"] = "svg",
    }.ToFrozenDictionary();

    /// <summary>
    /// Detects the MIME content type of an image from its binary data.
    /// </summary>
    /// <param name="data">Raw image bytes (at least first 12 bytes needed).</param>
    /// <returns>MIME type string (e.g., "image/jpeg"). Returns "application/octet-stream" if unknown.</returns>
    public static string GuessContentType(ReadOnlySpan<byte> data)
    {
        if (data.Length < 4)
            return "application/octet-stream";

        if (data.Length >= 8 && data[..8].SequenceEqual(PngSignature))
            return "image/png";
        if (data[..3].SequenceEqual(JpegSignature))
            return "image/jpeg";
        if (data.Length >= 6 && (data[..6].SequenceEqual(Gif87Signature) || data[..6].SequenceEqual(Gif89Signature)))
            return "image/gif";
        if (data[..2].SequenceEqual(BmpSignature))
            return "image/bmp";
        if (data[..4].SequenceEqual(TiffLeSignature) || data[..4].SequenceEqual(TiffBeSignature))
            return "image/tiff";
        if (data[..4].SequenceEqual(WmfSignature))
            return "image/x-wmf";

        // EMF: check for EMR_HEADER record type (1) and reasonable size
        if (data.Length >= 44 && data[..4].SequenceEqual(EmfSignature))
        {
            uint recordSize = BinaryPrimitives.ReadUInt32LittleEndian(data[4..]);
            if (recordSize >= 88)
                return "image/x-emf";
        }

        return "application/octet-stream";
    }

    /// <summary>
    /// Guesses the file extension for image data based on its content.
    /// </summary>
    /// <param name="data">Raw image bytes.</param>
    /// <returns>Extension string without dot (e.g., "jpeg", "png"). Returns "bin" if unknown.</returns>
    public static string GuessExtension(ReadOnlySpan<byte> data)
    {
        string contentType = GuessContentType(data);
        return ContentTypeExtensions.GetValueOrDefault(contentType, "bin");
    }

    /// <summary>
    /// Parses image dimensions from binary header data.
    /// </summary>
    /// <param name="data">Raw image bytes.</param>
    /// <returns>Tuple of (width, height) in pixels. Returns (0, 0) if format is unrecognized.</returns>
    public static (int Width, int Height) GetImageDimensions(ReadOnlySpan<byte> data)
    {
        if (data.Length < 4)
            return (0, 0);

        if (data.Length >= 8 && data[..8].SequenceEqual(PngSignature))
            return GetPngDimensions(data);
        if (data[..3].SequenceEqual(JpegSignature))
            return GetJpegDimensions(data);
        if (data.Length >= 6 && (data[..6].SequenceEqual(Gif87Signature) || data[..6].SequenceEqual(Gif89Signature)))
            return GetGifDimensions(data);
        if (data[..2].SequenceEqual(BmpSignature))
            return GetBmpDimensions(data);
        if (data[..4].SequenceEqual(TiffLeSignature) || data[..4].SequenceEqual(TiffBeSignature))
            return GetTiffDimensions(data);

        return (0, 0);
    }

    /// <summary>
    /// Parses PNG IHDR chunk for width and height.
    /// </summary>
    /// <param name="data">Raw PNG bytes.</param>
    /// <returns>Tuple of (width, height) in pixels.</returns>
    public static (int Width, int Height) GetPngDimensions(ReadOnlySpan<byte> data)
    {
        if (data.Length < 24)
            return (0, 0);

        // IHDR chunk: signature(8) + length(4) + 'IHDR'(4) + width(4) + height(4)
        int width = (int)BinaryPrimitives.ReadUInt32BigEndian(data[16..]);
        int height = (int)BinaryPrimitives.ReadUInt32BigEndian(data[20..]);
        return (width, height);
    }

    /// <summary>
    /// Parses JPEG SOF marker for width and height.
    /// </summary>
    /// <param name="data">Raw JPEG bytes.</param>
    /// <returns>Tuple of (width, height) in pixels.</returns>
    public static (int Width, int Height) GetJpegDimensions(ReadOnlySpan<byte> data)
    {
        int offset = 2; // Skip SOI marker
        int length = data.Length;

        while (offset < length - 1)
        {
            if (data[offset] != 0xFF)
            {
                offset++;
                continue;
            }

            byte marker = data[offset + 1];

            // Skip padding bytes
            if (marker == 0xFF)
            {
                offset++;
                continue;
            }

            // SOF markers (0xC0-0xCF, excluding 0xC4 DHT, 0xC8 JPG, 0xCC DAC)
            if (marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7
                       or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF)
            {
                if (offset + 9 < length)
                {
                    int height = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 5)..]);
                    int width = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 7)..]);
                    return (width, height);
                }
                return (0, 0);
            }

            // Skip other markers (read segment length)
            if (offset + 3 < length)
            {
                int segmentLength = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 2)..]);
                offset += 2 + segmentLength;
            }
            else
            {
                break;
            }
        }

        return (0, 0);
    }

    /// <summary>
    /// Parses GIF logical screen descriptor for width and height.
    /// </summary>
    /// <param name="data">Raw GIF bytes.</param>
    /// <returns>Tuple of (width, height) in pixels.</returns>
    public static (int Width, int Height) GetGifDimensions(ReadOnlySpan<byte> data)
    {
        if (data.Length < 10)
            return (0, 0);

        int width = BinaryPrimitives.ReadUInt16LittleEndian(data[6..]);
        int height = BinaryPrimitives.ReadUInt16LittleEndian(data[8..]);
        return (width, height);
    }

    /// <summary>
    /// Parses BMP info header for width and height.
    /// </summary>
    /// <param name="data">Raw BMP bytes.</param>
    /// <returns>Tuple of (width, height) in pixels. Height is always positive.</returns>
    public static (int Width, int Height) GetBmpDimensions(ReadOnlySpan<byte> data)
    {
        if (data.Length < 26)
            return (0, 0);

        int width = BinaryPrimitives.ReadInt32LittleEndian(data[18..]);
        int height = Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(data[22..])); // Height can be negative (top-down)
        return (width, height);
    }

    /// <summary>
    /// Parses TIFF IFD for ImageWidth and ImageLength tags.
    /// </summary>
    /// <param name="data">Raw TIFF bytes.</param>
    /// <returns>Tuple of (width, height) in pixels.</returns>
    public static (int Width, int Height) GetTiffDimensions(ReadOnlySpan<byte> data)
    {
        if (data.Length < 8)
            return (0, 0);

        // Determine byte order
        bool littleEndian = data[..2].SequenceEqual("II"u8);

        // Get offset to first IFD
        uint ifdOffset = littleEndian
            ? BinaryPrimitives.ReadUInt32LittleEndian(data[4..])
            : BinaryPrimitives.ReadUInt32BigEndian(data[4..]);

        if (ifdOffset + 2 > (uint)data.Length)
            return (0, 0);

        // Read number of IFD entries
        int ifdOff = (int)ifdOffset;
        ushort numEntries = littleEndian
            ? BinaryPrimitives.ReadUInt16LittleEndian(data[ifdOff..])
            : BinaryPrimitives.ReadUInt16BigEndian(data[ifdOff..]);

        int width = 0;
        int height = 0;
        int offset = ifdOff + 2;

        for (int i = 0; i < numEntries; i++)
        {
            if (offset + 12 > data.Length)
                break;

            ushort tag = littleEndian
                ? BinaryPrimitives.ReadUInt16LittleEndian(data[offset..])
                : BinaryPrimitives.ReadUInt16BigEndian(data[offset..]);

            ushort fieldType = littleEndian
                ? BinaryPrimitives.ReadUInt16LittleEndian(data[(offset + 2)..])
                : BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 2)..]);

            int valueOffset = offset + 8;
            int value;

            if (fieldType == 3) // SHORT
            {
                value = littleEndian
                    ? BinaryPrimitives.ReadUInt16LittleEndian(data[valueOffset..])
                    : BinaryPrimitives.ReadUInt16BigEndian(data[valueOffset..]);
            }
            else if (fieldType == 4) // LONG
            {
                value = (int)(littleEndian
                    ? BinaryPrimitives.ReadUInt32LittleEndian(data[valueOffset..])
                    : BinaryPrimitives.ReadUInt32BigEndian(data[valueOffset..]));
            }
            else
            {
                value = 0;
            }

            if (tag == 256) // ImageWidth
                width = value;
            else if (tag == 257) // ImageLength (height)
                height = value;

            if (width != 0 && height != 0)
                break;

            offset += 12;
        }

        return (width, height);
    }
}
