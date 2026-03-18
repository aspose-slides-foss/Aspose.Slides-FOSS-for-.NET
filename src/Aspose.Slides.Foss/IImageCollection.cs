namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of <see cref="IPPImage"/> objects.
/// </summary>
public abstract class IImageCollection
{
    /// <summary>
    /// Gets the collection as a list.
    /// </summary>
    public abstract IList<IPPImage> AsICollection { get; }

    /// <summary>
    /// Gets the collection as an enumerable.
    /// </summary>
    public abstract IEnumerable<IPPImage> AsIEnumerable { get; }

    /// <summary>
    /// Adds an image from an <see cref="IImage"/> instance.
    /// Returns an existing <see cref="IPPImage"/> if the data is a duplicate.
    /// </summary>
    /// <param name="image">The image to add.</param>
    /// <returns>The added or existing <see cref="IPPImage"/>.</returns>
    public abstract IPPImage AddImage(IImage image);

    /// <summary>
    /// Adds an image from a stream.
    /// Returns an existing <see cref="IPPImage"/> if the data is a duplicate.
    /// </summary>
    /// <param name="stream">The stream containing image data.</param>
    /// <returns>The added or existing <see cref="IPPImage"/>.</returns>
    public abstract IPPImage AddImage(Stream stream);

    /// <summary>
    /// Adds an image from raw byte data.
    /// Returns an existing <see cref="IPPImage"/> if the data is a duplicate.
    /// </summary>
    /// <param name="buffer">The image data as a byte array.</param>
    /// <returns>The added or existing <see cref="IPPImage"/>.</returns>
    public abstract IPPImage AddImage(byte[] buffer);

    /// <summary>
    /// Gets the number of images in the collection.
    /// </summary>
    public abstract int Count { get; }

    /// <summary>
    /// Gets an image by index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The <see cref="IPPImage"/> at the specified index.</returns>
    public abstract IPPImage this[int index] { get; }
}
