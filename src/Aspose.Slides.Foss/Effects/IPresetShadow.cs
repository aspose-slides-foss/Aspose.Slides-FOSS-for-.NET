namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a Preset Shadow effect.
/// </summary>
public interface IPresetShadow : IImageTransformOperation
{
    /// <summary>
    /// Gets or sets the direction of the shadow, in degrees.
    /// </summary>
    float Direction { get; set; }

    /// <summary>
    /// Gets or sets the distance of the shadow from the object, in points.
    /// </summary>
    double Distance { get; set; }

    /// <summary>
    /// Gets the color of the shadow. Read-only <see cref="IColorFormat"/>.
    /// </summary>
    IColorFormat ShadowColor { get; }

    /// <summary>
    /// Gets or sets the preset shadow type.
    /// </summary>
    PresetShadowType Preset { get; set; }

    /// <summary>
    /// Gets the base <see cref="IImageTransformOperation"/> interface. Read-only.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
