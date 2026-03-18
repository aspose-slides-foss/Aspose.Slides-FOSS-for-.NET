using Aspose.Slides.Foss.Drawing;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a color format used in presentation elements.
/// </summary>
public interface IColorFormat
{
    /// <summary>
    /// Gets or sets the color type.
    /// </summary>
    ColorType ColorType { get; set; }

    /// <summary>
    /// Gets or sets the RGB color value.
    /// </summary>
    Color? Color { get; set; }

    /// <summary>
    /// Gets or sets the preset color.
    /// </summary>
    PresetColor PresetColor { get; set; }

    /// <summary>
    /// Gets or sets the scheme color.
    /// </summary>
    SchemeColor SchemeColor { get; set; }

    /// <summary>
    /// Gets or sets the red component (0-255).
    /// </summary>
    int R { get; set; }

    /// <summary>
    /// Gets or sets the green component (0-255).
    /// </summary>
    int G { get; set; }

    /// <summary>
    /// Gets or sets the blue component (0-255).
    /// </summary>
    int B { get; set; }

    /// <summary>
    /// Gets or sets the red component as a float (0.0-1.0).
    /// </summary>
    float FloatR { get; set; }

    /// <summary>
    /// Gets or sets the green component as a float (0.0-1.0).
    /// </summary>
    float FloatG { get; set; }

    /// <summary>
    /// Gets or sets the blue component as a float (0.0-1.0).
    /// </summary>
    float FloatB { get; set; }

    /// <summary>
    /// Gets or sets the hue component (0-360).
    /// </summary>
    float Hue { get; set; }

    /// <summary>
    /// Gets or sets the saturation component (0-100).
    /// </summary>
    float Saturation { get; set; }

    /// <summary>
    /// Gets or sets the luminance component (0-100).
    /// </summary>
    float Luminance { get; set; }
}
