namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the geometric frame properties of a shape.
/// </summary>
public sealed class ShapeFrame : IShapeFrame, ICloneable
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ShapeFrame"/> class.
    /// </summary>
    /// <param name="x">The x-coordinate in points.</param>
    /// <param name="y">The y-coordinate in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <param name="flipH">Whether the shape is flipped horizontally.</param>
    /// <param name="flipV">Whether the shape is flipped vertically.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    public ShapeFrame(float x, float y, float width, float height, NullableBool flipH, NullableBool flipV, float rotation)
    {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        FlipH = flipH;
        FlipV = flipV;
        Rotation = rotation;
    }

    /// <inheritdoc/>
    public float X { get; }

    /// <inheritdoc/>
    public float Y { get; }

    /// <inheritdoc/>
    public float Width { get; }

    /// <inheritdoc/>
    public float Height { get; }

    /// <inheritdoc/>
    public NullableBool FlipH { get; }

    /// <inheritdoc/>
    public NullableBool FlipV { get; }

    /// <inheritdoc/>
    public float Rotation { get; }

    /// <inheritdoc/>
    public float CenterX => X + Width / 2;

    /// <inheritdoc/>
    public float CenterY => Y + Height / 2;

    /// <inheritdoc/>
    public object? Rectangle => (X, Y, Width, Height);

    /// <summary>
    /// Creates a deep copy of this <see cref="ShapeFrame"/>.
    /// </summary>
    /// <returns>A new object with the same property values.</returns>
    public object Clone() =>
        new ShapeFrame(X, Y, Width, Height, FlipH, FlipV, Rotation);

    /// <inheritdoc/>
    public IShapeFrame CloneT() =>
        new ShapeFrame(X, Y, Width, Height, FlipH, FlipV, Rotation);

    /// <summary>
    /// Determines whether the specified object is equal to this <see cref="ShapeFrame"/>.
    /// </summary>
    /// <param name="obj">The object to compare with.</param>
    /// <returns><c>true</c> if all seven fields are equal; otherwise, <c>false</c>.</returns>
    public override bool Equals(object? obj)
    {
        if (obj is not ShapeFrame other)
            return false;

        return X == other.X
            && Y == other.Y
            && Width == other.Width
            && Height == other.Height
            && FlipH == other.FlipH
            && FlipV == other.FlipV
            && Rotation == other.Rotation;
    }

    /// <summary>
    /// Returns a hash code for this instance.
    /// </summary>
    public override int GetHashCode() =>
        HashCode.Combine(X, Y, Width, Height, FlipH, FlipV, Rotation);
}
