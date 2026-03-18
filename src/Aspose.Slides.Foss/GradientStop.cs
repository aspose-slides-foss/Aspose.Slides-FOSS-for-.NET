using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single gradient stop within a gradient fill.
/// </summary>
public sealed class GradientStop : PVIObject, IGradientStop
{
    private XElement _gsElement = null!;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state for this gradient stop.
    /// </summary>
    internal void InitInternal(XElement gsElement, SlidePart slidePart, IBaseSlide? parentSlide)
    {
        _gsElement = gsElement;
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
    public float Position
    {
        get
        {
            var pos = _gsElement.Attribute("pos")?.Value;
            return pos is not null && int.TryParse(pos, out var v) ? v / 100000f : 0f;
        }
        set => _gsElement.SetAttributeValue("pos", (int)Math.Round(value * 100000));
    }

    /// <inheritdoc/>
    public IColorFormat Color
    {
        get
        {
            var cf = new ColorFormat();
            cf.InitInternal(_gsElement, _parentSlide, _slidePart);
            return cf;
        }
    }
}
