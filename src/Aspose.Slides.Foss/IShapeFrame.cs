namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the geometric frame properties of a shape.
/// </summary>
public interface IShapeFrame
{
    /// <summary>
    /// Gets the x-coordinate of the upper-left corner, in points.
    /// </summary>
    float X { get; }

    /// <summary>
    /// Gets the y-coordinate of the upper-left corner, in points.
    /// </summary>
    float Y { get; }

    /// <summary>
    /// Gets the width, in points.
    /// </summary>
    float Width { get; }

    /// <summary>
    /// Gets the height, in points.
    /// </summary>
    float Height { get; }

    /// <summary>
    /// Gets whether the shape is flipped horizontally.
    /// </summary>
    NullableBool FlipH { get; }

    /// <summary>
    /// Gets whether the shape is flipped vertically.
    /// </summary>
    NullableBool FlipV { get; }

    /// <summary>
    /// Gets the rotation angle in degrees.
    /// Positive values indicate clockwise rotation; negative values indicate counterclockwise.
    /// </summary>
    float Rotation { get; }

    /// <summary>
    /// Gets the x-coordinate of the frame's center, in points.
    /// </summary>
    float CenterX { get; }

    /// <summary>
    /// Gets the y-coordinate of the frame's center, in points.
    /// </summary>
    float CenterY { get; }

    /// <summary>
    /// Gets the bounding rectangle of the frame.
    /// </summary>
    object? Rectangle { get; }

    /// <summary>
    /// Creates a deep copy of this <see cref="IShapeFrame"/>.
    /// </summary>
    /// <returns>A new <see cref="IShapeFrame"/> instance with the same property values.</returns>
    IShapeFrame CloneT();
}
