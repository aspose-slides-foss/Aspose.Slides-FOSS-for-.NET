namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the bevel (relief) properties of a shape's face.
/// </summary>
public interface IShapeBevel
{
    /// <summary>
    /// Gets or sets the bevel width in points.
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// Gets or sets the bevel height in points.
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// Gets or sets the bevel preset type.
    /// </summary>
    BevelPresetType BevelType { get; set; }
}
