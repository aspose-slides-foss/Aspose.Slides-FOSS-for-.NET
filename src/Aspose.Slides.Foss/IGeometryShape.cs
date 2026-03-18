namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a shape with geometric properties.
/// </summary>
public interface IGeometryShape : IShape
{
    /// <summary>
    /// Gets the shape style reference.
    /// </summary>
    IShapeStyle? ShapeStyle { get; }

    /// <summary>
    /// Gets or sets the shape type.
    /// </summary>
    ShapeType ShapeType { get; set; }

    /// <summary>
    /// Gets the collection of adjustment values for this shape.
    /// </summary>
    IAdjustValueCollection? Adjustments { get; }
}
