namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents an Outer Shadow effect.
/// </summary>
public interface IOuterShadow : IImageTransformOperation
{
    /// <summary>
    /// Gets or sets the blur radius, in points. Default value is 0 pt.
    /// </summary>
    double BlurRadius { get; set; }

    /// <summary>
    /// Gets or sets the direction of the shadow, in degrees. Default value is 0.
    /// </summary>
    float Direction { get; set; }

    /// <summary>
    /// Gets or sets the distance of the shadow from the object, in points. Default value is 0 pt.
    /// </summary>
    double Distance { get; set; }

    /// <summary>
    /// Gets the color of the shadow. Read-only <see cref="IColorFormat"/>.
    /// </summary>
    IColorFormat ShadowColor { get; }

    /// <summary>
    /// Gets or sets the rectangle alignment.
    /// </summary>
    RectangleAlignment RectangleAlign { get; set; }

    /// <summary>
    /// Gets or sets the horizontal skew angle, in degrees. Default value is 0.
    /// </summary>
    double SkewHorizontal { get; set; }

    /// <summary>
    /// Gets or sets the vertical skew angle, in degrees. Default value is 0.
    /// </summary>
    double SkewVertical { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the shadow rotates together with the shape. Default value is true.
    /// </summary>
    bool RotateShadowWithShape { get; set; }

    /// <summary>
    /// Gets or sets the horizontal scaling factor, in percent. Negative scaling causes a flip. Default value is 100.
    /// </summary>
    double ScaleHorizontal { get; set; }

    /// <summary>
    /// Gets or sets the vertical scaling factor, in percent. Negative scaling causes a flip. Default value is 100.
    /// </summary>
    double ScaleVertical { get; set; }

    /// <summary>
    /// Gets the base <see cref="IImageTransformOperation"/> interface. Read-only.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
