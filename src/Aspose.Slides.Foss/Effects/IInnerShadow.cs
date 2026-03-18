namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents an inner shadow effect.
/// </summary>
public interface IInnerShadow : IImageTransformOperation
{
    /// <summary>
    /// Gets or sets the blur radius.
    /// </summary>
    double BlurRadius { get; set; }

    /// <summary>
    /// Gets or sets the direction of shadow.
    /// </summary>
    float Direction { get; set; }

    /// <summary>
    /// Gets or sets the distance of shadow.
    /// </summary>
    double Distance { get; set; }

    /// <summary>
    /// Gets the color of shadow. Read-only <see cref="IColorFormat"/>.
    /// </summary>
    IColorFormat ShadowColor { get; }

    /// <summary>
    /// Gets the base <see cref="IImageTransformOperation"/> interface. Read-only.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
