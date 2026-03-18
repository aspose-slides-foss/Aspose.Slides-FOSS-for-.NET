namespace Aspose.Slides.Foss;

/// <summary>
/// Represents any component that belongs to a slide.
/// </summary>
public abstract class ISlideComponent : IPresentationComponent
{
    /// <summary>
    /// Gets the owning base slide.
    /// </summary>
    public abstract IBaseSlide? Slide { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="IPresentationComponent"/>.
    /// </summary>
    public abstract IPresentationComponent AsIPresentationComponent { get; }
}
