using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a soft edge effect. The edges of the shape are blurred, while the fill is not affected.
/// </summary>
public sealed class SoftEdge : ISoftEdge
{
    private const float EmuPerPoint = 12700f;

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="SoftEdge"/> class.
    /// </summary>
    public SoftEdge()
    {
    }

    internal SoftEdge(XElement element)
    {
        _element = element;
    }

    /// <summary>
    /// Initializes the soft edge effect with its XML element, slide part, and parent slide.
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
    public IImageTransformOperation AsIImageTransformOperation => this;
}
