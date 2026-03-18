using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Provides static factory methods for creating <see cref="Image"/> instances.
/// </summary>
public static class Images
{
    /// <summary>
    /// Creates an <see cref="Image"/> from a file on disk.
    /// </summary>
    /// <param name="filename">The path to the image file.</param>
    /// <returns>A new <see cref="Image"/> initialized with the file data.</returns>
    public static Image FromFile(string filename)
    {
        var data = File.ReadAllBytes(filename);
        var contentType = ImageHelpers.GuessContentType(data);
        var image = new Image();
        image.InitInternal(data, contentType);
        return image;
    }

    /// <summary>
    /// Creates an <see cref="Image"/> from a stream.
    /// </summary>
    /// <param name="stream">The stream containing image data.</param>
    /// <returns>A new <see cref="Image"/> initialized with the stream data.</returns>
    public static Image FromStream(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var data = ms.ToArray();
        var contentType = ImageHelpers.GuessContentType(data);
        var image = new Image();
        image.InitInternal(data, contentType);
        return image;
    }
}
