namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a presentation-embedded image stored in an OPC package part.
/// </summary>
public abstract class IPPImage
{
    /// <summary>
    /// Gets a copy of the image data as a byte array.
    /// </summary>
    public abstract byte[] BinaryData { get; }

    /// <summary>
    /// Gets a copy of the image as an <see cref="IImage"/> instance.
    /// </summary>
    public abstract IImage Image { get; }

    /// <summary>
    /// Gets the MIME content type of the image.
    /// </summary>
    public abstract string ContentType { get; }

    /// <summary>
    /// Gets the image width in pixels.
    /// </summary>
    public abstract int Width { get; }

    /// <summary>
    /// Gets the image height in pixels.
    /// </summary>
    public abstract int Height { get; }

    /// <summary>
    /// Gets the X-offset of the image.
    /// </summary>
    public abstract int X { get; }

    /// <summary>
    /// Gets the Y-offset of the image.
    /// </summary>
    public abstract int Y { get; }

    /// <summary>
    /// Replaces this image with the specified raw byte data.
    /// </summary>
    /// <param name="newImageData">The new image data as a byte array.</param>
    public abstract void ReplaceImage(byte[] newImageData);

    /// <summary>
    /// Replaces this image with data from the specified <see cref="IImage"/>.
    /// </summary>
    /// <param name="newImage">The source image.</param>
    public abstract void ReplaceImage(IImage newImage);

    /// <summary>
    /// Replaces this image with data from another <see cref="IPPImage"/>.
    /// </summary>
    /// <param name="newImage">The source presentation image.</param>
    public abstract void ReplaceImage(IPPImage newImage);
}
