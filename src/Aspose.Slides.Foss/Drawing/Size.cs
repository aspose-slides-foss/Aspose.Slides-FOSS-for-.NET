namespace Aspose.Slides.Foss.Drawing;

/// <summary>
/// Represents a 2D size with integer dimensions.
/// </summary>
public sealed class Size
{
    /// <summary>
    /// Gets or sets the width.
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// Gets or sets the height.
    /// </summary>
    public int Height { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="Size"/> with the specified dimensions.
    /// </summary>
    public Size(int width = 0, int height = 0)
    {
        Width = width;
        Height = height;
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj) =>
        obj is Size other && Width == other.Width && Height == other.Height;

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Width, Height);

    /// <inheritdoc/>
    public override string ToString() => $"Size(width={Width}, height={Height})";
}
