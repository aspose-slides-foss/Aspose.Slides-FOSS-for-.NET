namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a blur effect that is applied to the entire shape, including its fill.
/// All color channels, including alpha, are affected.
/// </summary>
public interface IBlur : IImageTransformOperation
{
    /// <summary>
    /// Returns or sets the blur radius in points. Read/write <see cref="float"/>.
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// Determines whether the bounds of the object should be grown as a result of the blurring.
    /// <c>true</c> indicates the bounds are grown; <c>false</c> indicates they are not.
    /// Read/write <see cref="bool"/>.
    /// </summary>
    bool Grow { get; set; }

    /// <summary>
    /// Gets the parent slide. May be <c>null</c> if not associated with a slide.
    /// </summary>
    IBaseSlide? Slide { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="IPresentationComponent"/>.
    /// </summary>
    IPresentationComponent? AsIPresentationComponent { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="IImageTransformOperation"/>.
    /// </summary>
    IImageTransformOperation AsIImageTransformOperation { get; }
}
