using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Concrete image wrapper holding raw bytes and metadata.
/// </summary>
public sealed class Image : IImage, IDisposable
{
    private byte[] _data = Array.Empty<byte>();
    private string _contentType = string.Empty;
    private int _width;
    private int _height;

    /// <summary>
    /// Initializes internal state with image data and content type.
    /// </summary>
    internal void InitInternal(byte[] data, string contentType)
    {
        _data = data;
        _contentType = contentType;
        (_width, _height) = ImageHelpers.GetImageDimensions(data);
    }

    /// <summary>
    /// Gets the raw image data.
    /// </summary>
    internal byte[] Data => _data;

    /// <inheritdoc/>
    public override Size Size => new(_width, _height);

    /// <inheritdoc/>
    public override int Width => _width;

    /// <inheritdoc/>
    public override int Height => _height;

    /// <inheritdoc/>
    public override void Save(string filename)
    {
        File.WriteAllBytes(filename, _data);
    }

    /// <inheritdoc/>
    public override void Save(string filename, string format)
    {
        File.WriteAllBytes(filename, _data);
    }

    /// <inheritdoc/>
    public override void Save(Stream stream, string format)
    {
        stream.Write(_data, 0, _data.Length);
    }

    /// <inheritdoc/>
    public override void Save(string filename, string format, int quality)
    {
        File.WriteAllBytes(filename, _data);
    }

    /// <inheritdoc/>
    public override void Save(Stream stream, string format, int quality)
    {
        stream.Write(_data, 0, _data.Length);
    }

    /// <summary>
    /// Releases resources held by this image.
    /// </summary>
    public void Dispose()
    {
        // No unmanaged resources; provided for context-manager protocol parity.
    }
}
