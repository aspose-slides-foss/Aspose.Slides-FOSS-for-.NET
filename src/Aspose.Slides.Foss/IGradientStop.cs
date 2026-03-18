namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single stop in a gradient fill.
/// </summary>
public interface IGradientStop
{
    /// <summary>
    /// Gets or sets the stop position (0..1).
    /// </summary>
    float Position { get; set; }

    /// <summary>
    /// Gets the color at this gradient stop.
    /// </summary>
    IColorFormat Color { get; }
}
