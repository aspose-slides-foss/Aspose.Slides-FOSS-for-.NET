namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single adjustment value for a geometry shape.
/// </summary>
public interface IAdjustValue
{
    /// <summary>
    /// Gets the name of the adjustment value.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets or sets the raw integer value from the formula attribute.
    /// </summary>
    int RawValue { get; set; }

    /// <summary>
    /// Gets or sets the value expressed in degrees (raw value / 60000).
    /// </summary>
    float AngleValue { get; set; }
}
