using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Concrete presentation-embedded image backed by an OPC package part.
/// </summary>
public sealed class PPImage : IPPImage
{
    private OpcPackage? _package;
    private string _partName = string.Empty;
    private byte[] _imageData = Array.Empty<byte>();
    private string _contentType = string.Empty;
    private int _width;
    private int _height;

    /// <summary>
    /// Initializes internal state with package, part name, data, and content type.
    /// </summary>
    internal void InitInternal(OpcPackage package, string partName, byte[] imageData, string contentType)
    {
        _package = package;
        _partName = partName;
        _imageData = imageData;
        _contentType = contentType;
        (_width, _height) = ImageHelpers.GetImageDimensions(imageData);
    }

    /// <summary>
    /// Gets the OPC part name for this image (e.g., "ppt/media/image1.jpg").
    /// </summary>
    internal string PartName => _partName;

    /// <summary>
    /// Gets the raw image data. For internal use by other image-related classes.
    /// </summary>
    internal byte[] ImageData => _imageData;

    /// <inheritdoc/>
    public override byte[] BinaryData => (byte[])_imageData.Clone();

    /// <inheritdoc/>
    public override IImage Image
    {
        get
        {
            var img = new Image();
            img.InitInternal((byte[])_imageData.Clone(), _contentType);
            return img;
        }
    }

    /// <inheritdoc/>
    public override string ContentType => _contentType;

    /// <inheritdoc/>
    public override int Width => _width;

    /// <inheritdoc/>
    public override int Height => _height;

    /// <inheritdoc/>
    public override int X => 0;

    /// <inheritdoc/>
    public override int Y => 0;

    /// <inheritdoc/>
    public override void ReplaceImage(byte[] newImageData)
    {
        ReplaceImageCore(newImageData);
    }

    /// <inheritdoc/>
    public override void ReplaceImage(IImage newImage)
    {
        if (newImage is Image img)
        {
            ReplaceImageCore(img.Data);
        }
        else
        {
            throw new ArgumentException("Unsupported IImage implementation.", nameof(newImage));
        }
    }

    /// <inheritdoc/>
    public override void ReplaceImage(IPPImage newImage)
    {
        if (newImage is PPImage ppImg)
        {
            ReplaceImageCore(ppImg._imageData);
        }
        else
        {
            throw new ArgumentException("Unsupported IPPImage implementation.", nameof(newImage));
        }
    }

    private void ReplaceImageCore(byte[] newData)
    {
        _imageData = newData;
        _contentType = ImageHelpers.GuessContentType(newData);
        (_width, _height) = ImageHelpers.GetImageDimensions(newData);
        _package?.SetPart(_partName, newData);
    }
}
