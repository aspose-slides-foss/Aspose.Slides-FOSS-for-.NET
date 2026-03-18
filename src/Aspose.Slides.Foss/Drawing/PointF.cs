namespace Aspose.Slides.Foss.Drawing;

/// <summary>
/// Represents a 2D point with float coordinates.
/// </summary>
public sealed class PointF
{
    /// <summary>
    /// Gets or sets the X coordinate.
    /// </summary>
    public float X { get; set; }

    /// <summary>
    /// Gets or sets the Y coordinate.
    /// </summary>
    public float Y { get; set; }

    /// <summary>
    /// Initializes a new instance of <see cref="PointF"/> with the specified coordinates.
    /// </summary>
    public PointF(float x = 0f, float y = 0f)
    {
        X = x;
        Y = y;
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj) =>
        obj is PointF other && X == other.X && Y == other.Y;

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(X, Y);

    /// <inheritdoc/>
    public override string ToString() => $"PointF(x={X}, y={Y})";
}
