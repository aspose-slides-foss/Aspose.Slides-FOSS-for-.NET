namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a glow effect, in which a color blurred outline is added outside the edges of the object.
/// </summary>
public interface IGlow : IImageTransformOperation
{
    /// <summary>
    /// Returns or sets the glow radius in points. Read/write <see cref="float"/>.
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// Returns the color format of the glow effect. Read-only <see cref="IColorFormat"/>.
    /// </summary>
    IColorFormat Color { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="IImageTransformOperation"/>.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
