using System.IO.Compression;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// </summary>
internal static class TestHelpers
{
    /// <summary>
    /// Generates a minimal valid 1×1 PNG with the given RGB colour.
    /// </summary>
    internal static byte[] CreateTestPng(byte r = 255, byte g = 0, byte b = 0)
    {
        static byte[] Chunk(byte[] chunkType, byte[] data)
        {
            var combined = new byte[chunkType.Length + data.Length];
            Buffer.BlockCopy(chunkType, 0, combined, 0, chunkType.Length);
            Buffer.BlockCopy(data, 0, combined, chunkType.Length, data.Length);

            var crc = ComputeCrc32(combined);
            var lengthBytes = ToBigEndianBytes((uint)data.Length);
            var crcBytes = ToBigEndianBytes(crc);

            var result = new byte[4 + combined.Length + 4];
            Buffer.BlockCopy(lengthBytes, 0, result, 0, 4);
            Buffer.BlockCopy(combined, 0, result, 4, combined.Length);
            Buffer.BlockCopy(crcBytes, 0, result, 4 + combined.Length, 4);
            return result;
        }

        byte[] header = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        // IHDR: width=1, height=1, bit_depth=8, color_type=2 (RGB), compression=0, filter=0, interlace=0
        using var ihdrStream = new MemoryStream();
        using (var w = new BinaryWriter(ihdrStream))
        {
            w.Write(ToBigEndianBytes(1u)); // width
            w.Write(ToBigEndianBytes(1u)); // height
            w.Write((byte)8);  // bit depth
            w.Write((byte)2);  // color type (RGB)
            w.Write((byte)0);  // compression
            w.Write((byte)0);  // filter method
            w.Write((byte)0);  // interlace
        }
        var ihdr = ihdrStream.ToArray();

        // Raw scanline: filter byte (0) + RGB
        byte[] raw = [0, r, g, b];

        // Compress with zlib (DeflateStream wrapped with zlib header/trailer)
        var idat = ZlibCompress(raw);

        var ihdrChunk = Chunk("IHDR"u8.ToArray(), ihdr);
        var idatChunk = Chunk("IDAT"u8.ToArray(), idat);
        var iendChunk = Chunk("IEND"u8.ToArray(), []);

        using var result = new MemoryStream();
        result.Write(header);
        result.Write(ihdrChunk);
        result.Write(idatChunk);
        result.Write(iendChunk);
        return result.ToArray();
    }

    /// <summary>
    /// Saves a <see cref="Presentation"/> to a temp file, disposes it, and reopens from that file.
    /// Creates a temporary PPTX file for testing.
    /// </summary>
    internal static Presentation SaveAndReopen(Presentation pres, string tempDir)
    {
        var path = Path.Combine(tempDir, "roundtrip.pptx");
        pres.Save(path, SaveFormat.Pptx);
        pres.Dispose();
        return new Presentation(path);
    }

    private static byte[] ToBigEndianBytes(uint value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return bytes;
    }

    private static byte[] ZlibCompress(byte[] data)
    {
        using var output = new MemoryStream();
        // zlib header: CMF=0x78 (deflate, window size 32K), FLG=0x9C
        output.WriteByte(0x78);
        output.WriteByte(0x9C);

        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            deflate.Write(data, 0, data.Length);
        }

        // Adler-32 checksum
        var adler = ComputeAdler32(data);
        var adlerBytes = ToBigEndianBytes(adler);
        output.Write(adlerBytes, 0, 4);

        return output.ToArray();
    }

    private static uint ComputeAdler32(byte[] data)
    {
        const uint mod = 65521;
        uint a = 1, b = 0;
        foreach (var d in data)
        {
            a = (a + d) % mod;
            b = (b + a) % mod;
        }
        return (b << 16) | a;
    }

    private static uint ComputeCrc32(byte[] data)
    {
        // Standard CRC-32
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (var i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (0xEDB88320 & ~((crc & 1) - 1));
        }
        return crc ^ 0xFFFFFFFF;
    }
}
