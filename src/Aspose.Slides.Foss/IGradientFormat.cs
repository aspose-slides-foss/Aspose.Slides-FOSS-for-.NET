namespace Aspose.Slides.Foss;

/// <summary>
/// Represents gradient fill formatting properties.
/// </summary>
public interface IGradientFormat : IFillParamSource
{
    /// <summary>
    /// Gets or sets the tile flip mode.
    /// </summary>
    TileFlip TileFlip { get; set; }

    /// <summary>
    /// Gets or sets the gradient direction.
    /// </summary>
    GradientDirection GradientDirection { get; set; }

    /// <summary>
    /// Gets or sets the gradient shape.
    /// </summary>
    GradientShape GradientShape { get; set; }

    /// <summary>
    /// Gets the collection of gradient stops.
    /// </summary>
    IGradientStopCollection GradientStops { get; }

    /// <summary>
    /// Gets or sets the linear gradient angle in degrees.
    /// </summary>
    float LinearGradientAngle { get; set; }

    /// <summary>
    /// Gets or sets whether the linear gradient is scaled.
    /// </summary>
    NullableBool LinearGradientScaled { get; set; }
}
