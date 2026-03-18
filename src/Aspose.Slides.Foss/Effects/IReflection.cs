namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a reflection effect.
/// </summary>
public interface IReflection : IImageTransformOperation
{
    /// <summary>
    /// Gets or sets the start position (along the alpha gradient ramp) of the start alpha value, in percent.
    /// </summary>
    float StartPosAlpha { get; set; }

    /// <summary>
    /// Gets or sets the end position (along the alpha gradient ramp) of the end alpha value, in percent.
    /// </summary>
    float EndPosAlpha { get; set; }

    /// <summary>
    /// Gets or sets the direction to offset the reflection, in degrees.
    /// </summary>
    float FadeDirection { get; set; }

    /// <summary>
    /// Gets or sets the starting reflection opacity, in percent.
    /// </summary>
    float StartReflectionOpacity { get; set; }

    /// <summary>
    /// Gets or sets the end reflection opacity, in percent.
    /// </summary>
    float EndReflectionOpacity { get; set; }

    /// <summary>
    /// Gets or sets the blur radius, in points.
    /// </summary>
    double BlurRadius { get; set; }

    /// <summary>
    /// Gets or sets the direction of the reflection, in degrees.
    /// </summary>
    float Direction { get; set; }

    /// <summary>
    /// Gets or sets the distance of the reflection, in points.
    /// </summary>
    double Distance { get; set; }

    /// <summary>
    /// Gets or sets the rectangle alignment.
    /// </summary>
    RectangleAlignment RectangleAlign { get; set; }

    /// <summary>
    /// Gets or sets the horizontal skew angle, in degrees.
    /// </summary>
    double SkewHorizontal { get; set; }

    /// <summary>
    /// Gets or sets the vertical skew angle, in degrees.
    /// </summary>
    double SkewVertical { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the reflection rotates with the shape.
    /// </summary>
    bool RotateShadowWithShape { get; set; }

    /// <summary>
    /// Gets or sets the horizontal scaling factor, in percent. Negative scaling causes a flip.
    /// </summary>
    double ScaleHorizontal { get; set; }

    /// <summary>
    /// Gets or sets the vertical scaling factor, in percent. Negative scaling causes a flip.
    /// </summary>
    double ScaleVertical { get; set; }

    /// <summary>
    /// Gets the base <see cref="IImageTransformOperation"/> interface. Read-only.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
