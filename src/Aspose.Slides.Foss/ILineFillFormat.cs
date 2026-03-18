namespace Aspose.Slides.Foss;

/// <summary>
/// Represents properties for lines filling.
/// </summary>
public interface ILineFillFormat
{
    /// <summary>
    /// Returns or sets the fill type. Read/write.
    /// </summary>
    FillType FillType { get; set; }

    /// <summary>
    /// Returns the color of a solid fill. Read-only.
    /// </summary>
    IColorFormat SolidFillColor { get; }

    /// <summary>
    /// Returns the gradient fill format. Read-only.
    /// </summary>
    IGradientFormat GradientFormat { get; }

    /// <summary>
    /// Returns the pattern fill format. Read-only.
    /// </summary>
    IPatternFormat PatternFormat { get; }

    /// <summary>
    /// Determines whether the fill should be rotated with a shape. Read/write.
    /// </summary>
    NullableBool RotateWithShape { get; set; }
}
