namespace Aspose.Slides.Foss;

/// <summary>
/// Represents fill formatting properties for a shape or text.
/// </summary>
public interface IFillFormat : IFillParamSource
{
    /// <summary>
    /// Gets or sets the fill type.
    /// </summary>
    FillType FillType { get; set; }

    /// <summary>
    /// Gets the solid fill color.
    /// </summary>
    IColorFormat SolidFillColor { get; }

    /// <summary>
    /// Gets the gradient fill format.
    /// </summary>
    IGradientFormat GradientFormat { get; }

    /// <summary>
    /// Gets the pattern fill format.
    /// </summary>
    IPatternFormat PatternFormat { get; }

    /// <summary>
    /// Gets the picture fill format.
    /// </summary>
    IPictureFillFormat PictureFillFormat { get; }

    /// <summary>
    /// Gets or sets whether the fill rotates with the shape.
    /// </summary>
    NullableBool RotateWithShape { get; set; }
}
