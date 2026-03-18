using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a glow effect, in which a color blurred outline is added outside the edges of the object.
/// </summary>
public sealed class Glow : IGlow
{
    private const float EmuPerPoint = 12700f;

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="Glow"/> class.
    /// </summary>
    public Glow()
    {
    }

    internal Glow(XElement element)
    {
        _element = element;
    }

    /// <summary>
    /// Initializes the glow effect with its XML element, slide part, and parent slide.
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
    public IColorFormat Color
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
    public IImageTransformOperation AsIImageTransformOperation => this;
}
