namespace Aspose.Slides.Foss;

/// <summary>
/// Concrete base class providing property-value-inheritance infrastructure.
/// </summary>
public class PVIObject : ISlideComponent
{
    /// <summary>
    /// The parent slide reference. Subclasses are responsible for setting this field.
    /// </summary>
    protected IBaseSlide? _parentSlide;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;
}
