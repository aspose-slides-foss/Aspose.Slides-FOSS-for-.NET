using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a blur effect that is applied to the entire shape, including its fill.
/// All color channels, including alpha, are affected.
/// </summary>
public sealed class Blur : ISlideComponent, IBlur, IImageTransformOperation
{
    private const float EmuPerPoint = 12700f;

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="Blur"/> class.
    /// </summary>
    public Blur()
    {
    }

    internal Blur(XElement element, IBaseSlide? parentSlide = null)
    {
        _element = element;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Initializes the blur effect with its XML element, slide part, and parent slide.
    /// </summary>
    internal void InitInternal(XElement element, SlidePart slidePart, IBaseSlide? parentSlide)
    {
        _element = element;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Persists changes to the underlying slide part.
    /// </summary>
    internal void Save()
    {
        _slidePart?.Save();
    }

    /// <inheritdoc/>
    public float Radius
    {
        get
        {
            var val = _element?.Attribute("rad")?.Value;
            return val is null ? 0f : int.Parse(val) / EmuPerPoint;
        }
        set
        {
            _element?.SetAttributeValue("rad", (int)MathF.Round(value * EmuPerPoint));
        }
    }

    /// <inheritdoc/>
    public bool Grow
    {
        get
        {
            var val = _element?.Attribute("grow")?.Value;
            return val is null || val == "1";
        }
        set
        {
            _element?.SetAttributeValue("grow", value ? "1" : "0");
        }
    }

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public IImageTransformOperation AsIImageTransformOperation => this;
}
