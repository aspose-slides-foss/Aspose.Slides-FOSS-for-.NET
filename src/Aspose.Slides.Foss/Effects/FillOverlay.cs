using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a Fill Overlay effect. A fill overlay may be used to specify
/// an additional fill for an object and blend the two fills together.
/// </summary>
public sealed class FillOverlay : ISlideComponent, IFillOverlay, IImageTransformOperation
{
    /// <summary>OOXML blend attribute value to <see cref="FillBlendMode"/> mapping.</summary>
    private static readonly Dictionary<string, FillBlendMode> BlendMap = new()
    {
        ["over"] = FillBlendMode.Overlay,
        ["mult"] = FillBlendMode.Multiply,
        ["screen"] = FillBlendMode.Screen,
        ["darken"] = FillBlendMode.Darken,
        ["lighten"] = FillBlendMode.Lighten,
    };

    /// <summary>Reverse mapping from <see cref="FillBlendMode"/> to OOXML attribute value.</summary>
    private static readonly Dictionary<FillBlendMode, string> BlendMapReverse = new()
    {
        [FillBlendMode.Overlay] = "over",
        [FillBlendMode.Multiply] = "mult",
        [FillBlendMode.Screen] = "screen",
        [FillBlendMode.Darken] = "darken",
        [FillBlendMode.Lighten] = "lighten",
    };

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="FillOverlay"/> class.
    /// </summary>
    public FillOverlay()
    {
    }

    internal FillOverlay(XElement element, IBaseSlide? parentSlide = null)
    {
        _element = element;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Initializes the fill overlay effect with its XML element, slide part, and parent slide.
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
    public IFillFormat FillFormat
    {
        get
        {
            var ff = new FillFormat();
            if (_element is not null)
                ff.InitInternal(_element, _parentSlide);
            return ff;
        }
    }

    /// <inheritdoc/>
    public FillBlendMode Blend
    {
        get
        {
            var val = _element?.Attribute("blend")?.Value;
            if (val is not null && BlendMap.TryGetValue(val, out var mode))
                return mode;
            return FillBlendMode.Overlay;
        }
        set
        {
            if (_element is not null && BlendMapReverse.TryGetValue(value, out var ooxmlVal))
                _element.SetAttributeValue("blend", ooxmlVal);
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
