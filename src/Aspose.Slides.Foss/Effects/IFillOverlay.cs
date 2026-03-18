namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a Fill Overlay effect. A fill overlay may be used to specify
/// an additional fill for an object and blend the two fills together.
/// </summary>
public interface IFillOverlay : IImageTransformOperation
{
    /// <summary>
    /// Gets the fill format. Read-only <see cref="IFillFormat"/>.
    /// </summary>
    IFillFormat FillFormat { get; }

    /// <summary>
    /// Gets or sets the fill blend mode. Read/write <see cref="FillBlendMode"/>.
    /// </summary>
    FillBlendMode Blend { get; set; }

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
