using Aspose.Slides.Foss.Drawing;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a raster or vector image.
/// </summary>
public abstract class IImage
{
    /// <summary>
    /// Gets the image dimensions as a <see cref="Drawing.Size"/>.
    /// </summary>
    public abstract Size Size { get; }

    /// <summary>
    /// Gets the image width in pixels.
    /// </summary>
    public abstract int Width { get; }

    /// <summary>
    /// Gets the image height in pixels.
    /// </summary>
    public abstract int Height { get; }

    /// <summary>
    /// Saves the image to the specified file path.
    /// </summary>
    /// <param name="filename">The destination file path.</param>
    public abstract void Save(string filename);

    /// <summary>
    /// Saves the image to the specified file path in the given format.
    /// </summary>
    /// <param name="filename">The destination file path.</param>
    /// <param name="format">The image format.</param>
    public abstract void Save(string filename, string format);

    /// <summary>
    /// Saves the image to the specified stream in the given format.
    /// </summary>
    /// <param name="stream">The destination stream.</param>
    /// <param name="format">The image format.</param>
    public abstract void Save(Stream stream, string format);

    /// <summary>
    /// Saves the image to the specified file path in the given format and quality.
    /// </summary>
    /// <param name="filename">The destination file path.</param>
    /// <param name="format">The image format.</param>
    /// <param name="quality">The image quality (0-100).</param>
    public abstract void Save(string filename, string format, int quality);

    /// <summary>
    /// Saves the image to the specified stream in the given format and quality.
    /// </summary>
    /// <param name="stream">The destination stream.</param>
    /// <param name="format">The image format.</param>
    /// <param name="quality">The image quality (0-100).</param>
    public abstract void Save(Stream stream, string format, int quality);
}
