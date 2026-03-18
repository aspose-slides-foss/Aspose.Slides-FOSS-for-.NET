namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a Soft Edge effect. The edges of the shape are blurred, while the fill is not affected.
/// </summary>
public interface ISoftEdge : IImageTransformOperation
{
    /// <summary>
    /// Specifies the radius of blur to apply to the edges, in points. Read/write <see cref="float"/>.
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="IImageTransformOperation"/>. Read-only.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
