namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Utility methods for image dimension extraction and content type guessing.
/// </summary>
internal static class ImageHelpers
{
    /// <summary>
    /// Extracts width and height from image binary data by inspecting file headers.
    /// Returns (0, 0) if the format is not recognized.
    /// </summary>
    internal static (int Width, int Height) GetImageDimensions(byte[] data)
    {
        if (data.Length < 8)
            return (0, 0);

        // PNG: 8-byte signature, then IHDR chunk with width/height at bytes 16-23
        if (data.Length >= 24 &&
            data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
        {
            int width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
            int height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
            return (width, height);
        }

        // JPEG: SOI marker 0xFF 0xD8, scan for SOFn markers
        if (data[0] == 0xFF && data[1] == 0xD8)
        {
            return GetJpegDimensions(data);
        }

        // GIF: "GIF87a" or "GIF89a"
        if (data.Length >= 10 &&
            data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
        {
            int width = data[6] | (data[7] << 8);
            int height = data[8] | (data[9] << 8);
            return (width, height);
        }

        // BMP: "BM" header
        if (data.Length >= 26 &&
            data[0] == 0x42 && data[1] == 0x4D)
        {
            int width = data[18] | (data[19] << 8) | (data[20] << 16) | (data[21] << 24);
            int height = data[22] | (data[23] << 8) | (data[24] << 16) | (data[25] << 24);
            if (height < 0) height = -height;
            return (width, height);
        }

        // TIFF: "II" (little-endian) or "MM" (big-endian)
        if (data.Length >= 8 &&
            ((data[0] == 0x49 && data[1] == 0x49) || (data[0] == 0x4D && data[1] == 0x4D)))
        {
            return GetTiffDimensions(data);
        }

        return (0, 0);
    }

    private static (int Width, int Height) GetJpegDimensions(byte[] data)
    {
        int i = 2;
        while (i + 1 < data.Length)
        {
            if (data[i] != 0xFF)
                break;

            byte marker = data[i + 1];

            // Skip filler bytes
            if (marker == 0xFF)
            {
                i++;
                continue;
            }

            // SOFn markers: 0xC0-0xC3, 0xC5-0xC7, 0xC9-0xCB, 0xCD-0xCF
            if (marker is >= 0xC0 and <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
            {
                if (i + 9 < data.Length)
                {
                    int height = (data[i + 5] << 8) | data[i + 6];
                    int width = (data[i + 7] << 8) | data[i + 8];
                    return (width, height);
                }
            }

            // Skip to next marker
            if (i + 3 < data.Length)
            {
                int segLen = (data[i + 2] << 8) | data[i + 3];
                i += 2 + segLen;
            }
            else
            {
                break;
            }
        }

        return (0, 0);
    }

    private static (int Width, int Height) GetTiffDimensions(byte[] data)
    {
        bool littleEndian = data[0] == 0x49;

        int ReadU16(int offset)
        {
            if (offset + 1 >= data.Length) return 0;
            return littleEndian
                ? data[offset] | (data[offset + 1] << 8)
                : (data[offset] << 8) | data[offset + 1];
        }

        int ReadU32(int offset)
        {
            if (offset + 3 >= data.Length) return 0;
            return littleEndian
                ? data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24)
                : (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
        }

        int ifdOffset = ReadU32(4);
        if (ifdOffset <= 0 || ifdOffset >= data.Length)
            return (0, 0);

        int entryCount = ReadU16(ifdOffset);
        int width = 0, height = 0;

        for (int e = 0; e < entryCount; e++)
        {
            int entryOffset = ifdOffset + 2 + e * 12;
            if (entryOffset + 12 > data.Length)
                break;

            int tag = ReadU16(entryOffset);
            int type = ReadU16(entryOffset + 2);
            int valueOffset = entryOffset + 8;

            int value = type == 3 ? ReadU16(valueOffset) : ReadU32(valueOffset);

            if (tag == 256) width = value;   // ImageWidth
            if (tag == 257) height = value;  // ImageLength

            if (width > 0 && height > 0)
                return (width, height);
        }

        return (width, height);
    }

    /// <summary>
    /// Guesses the MIME content type from image binary data by inspecting file signatures.
    /// </summary>
    internal static string GuessContentType(byte[] data)
    {
        if (data.Length >= 8 &&
            data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            return "image/png";

        if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xD8)
            return "image/jpeg";

        if (data.Length >= 6 &&
            data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
            return "image/gif";

        if (data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D)
            return "image/bmp";

        if (data.Length >= 4 &&
            ((data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00) ||
             (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A)))
            return "image/tiff";

        if (data.Length >= 4 &&
            data[0] == 0x52 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x46 &&
            data.Length >= 12 &&
            data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50)
            return "image/webp";

        // Check for SVG (XML-based)
        if (data.Length >= 5)
        {
            var head = System.Text.Encoding.UTF8.GetString(data, 0, Math.Min(data.Length, 256));
            if (head.Contains("<svg", StringComparison.OrdinalIgnoreCase))
                return "image/svg+xml";
        }

        // WMF
        if (data.Length >= 4 &&
            data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A)
            return "image/x-wmf";

        // EMF
        if (data.Length >= 4 &&
            data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00 &&
            data.Length >= 44 &&
            data[40] == 0x20 && data[41] == 0x45 && data[42] == 0x4D && data[43] == 0x46)
            return "image/x-emf";

        return "application/octet-stream";
    }

    /// <summary>
    /// Gets the file extension for a given MIME content type.
    /// </summary>
    internal static string GetExtensionForContentType(string contentType)
    {
        return contentType.ToLowerInvariant() switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpeg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/webp" => "webp",
            "image/svg+xml" => "svg",
            "image/x-wmf" => "wmf",
            "image/x-emf" => "emf",
            _ => "bin",
        };
    }
}
