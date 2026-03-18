namespace Aspose.Slides.Foss.Drawing;

/// <summary>
/// Represents a 2D size with float dimensions.
/// </summary>
public sealed class SizeF
{
    /// <summary>
    /// Gets or sets the width.
    /// </summary>
    public float Width { get; set; }

    /// <summary>
    /// Gets or sets the height.
    /// </summary>
    public float Height { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="SizeF"/> with the specified dimensions.
    /// </summary>
    public SizeF(float width = 0f, float height = 0f)
    {
        Width = width;
        Height = height;
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj) =>
        obj is SizeF other && Width == other.Width && Height == other.Height;

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Width, Height);

    /// <inheritdoc/>
    public override string ToString() => $"SizeF(width={Width}, height={Height})";
}
