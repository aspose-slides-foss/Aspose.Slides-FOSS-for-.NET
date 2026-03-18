using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents an inner shadow effect.
/// </summary>
public sealed class InnerShadow : ISlideComponent, IInnerShadow, IImageTransformOperation
{
    private const float EmuPerPoint = 12700f;

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="InnerShadow"/> class.
    /// </summary>
    public InnerShadow()
    {
    }

    internal InnerShadow(XElement element, IBaseSlide? parentSlide = null)
    {
        _element = element;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Initializes the inner shadow effect with its XML element, slide part, and parent slide.
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
    public double BlurRadius
    {
        get
        {
            var val = _element?.Attribute("blurRad")?.Value;
            return val is null ? 0d : int.Parse(val) / (double)EmuPerPoint;
        }
        set
        {
            _element?.SetAttributeValue("blurRad", (int)Math.Round(value * EmuPerPoint));
        }
    }

    /// <inheritdoc/>
    public float Direction
    {
        get
        {
            var val = _element?.Attribute("dir")?.Value;
            return val is null ? 0f : int.Parse(val) / 60000f;
        }
        set
        {
            _element?.SetAttributeValue("dir", (int)MathF.Round(value * 60000f));
        }
    }

    /// <inheritdoc/>
    public double Distance
    {
        get
        {
            var val = _element?.Attribute("dist")?.Value;
            return val is null ? 0d : int.Parse(val) / (double)EmuPerPoint;
        }
        set
        {
            _element?.SetAttributeValue("dist", (int)Math.Round(value * EmuPerPoint));
        }
    }

    /// <inheritdoc/>
    public IColorFormat ShadowColor
    {
        get
        {
            var colorFormat = new ColorFormat();
            if (_element is not null)
            {
                colorFormat.InitInternal(_element, null);
            }
            return colorFormat;
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
