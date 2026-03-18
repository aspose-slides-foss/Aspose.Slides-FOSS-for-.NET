namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a group shape that contains other shapes.
/// </summary>
public interface IGroupShape : IShape
{
    /// <summary>
    /// Gets the collection of shapes within this group.
    /// </summary>
    IShapeCollection Shapes { get; }
}
