using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents an outer shadow effect.
/// </summary>
public sealed class OuterShadow : ISlideComponent, IOuterShadow, IImageTransformOperation
{
    private const float EmuPerPoint = 12700f;

    private static readonly Dictionary<string, RectangleAlignment> AlgnMap = new()
    {
        ["tl"] = RectangleAlignment.TopLeft,
        ["t"] = RectangleAlignment.Top,
        ["tr"] = RectangleAlignment.TopRight,
        ["l"] = RectangleAlignment.Left,
        ["ctr"] = RectangleAlignment.Center,
        ["r"] = RectangleAlignment.Right,
        ["bl"] = RectangleAlignment.BottomLeft,
        ["b"] = RectangleAlignment.Bottom,
        ["br"] = RectangleAlignment.BottomRight,
    };

    private static readonly Dictionary<RectangleAlignment, string> AlgnMapRev =
        AlgnMap.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="OuterShadow"/> class.
    /// </summary>
    public OuterShadow()
    {
    }

    internal OuterShadow(XElement element, IBaseSlide? parentSlide = null)
    {
        _element = element;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Initializes the outer shadow effect with its XML element, slide part, and parent slide.
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
    public RectangleAlignment RectangleAlign
    {
        get
        {
            var val = _element?.Attribute("algn")?.Value;
            if (val is null) return RectangleAlignment.Bottom;
            return AlgnMap.GetValueOrDefault(val, RectangleAlignment.NotDefined);
        }
        set
        {
            if (value == RectangleAlignment.NotDefined)
            {
                _element?.Attribute("algn")?.Remove();
            }
            else if (AlgnMapRev.TryGetValue(value, out var ooxmlVal))
            {
                _element?.SetAttributeValue("algn", ooxmlVal);
            }
        }
    }

    /// <inheritdoc/>
    public double SkewHorizontal
    {
        get
        {
            var val = _element?.Attribute("kx")?.Value;
            return val is null ? 0d : int.Parse(val) / 60000d;
        }
        set
        {
            _element?.SetAttributeValue("kx", (int)Math.Round(value * 60000d));
        }
    }

    /// <inheritdoc/>
    public double SkewVertical
    {
        get
        {
            var val = _element?.Attribute("ky")?.Value;
            return val is null ? 0d : int.Parse(val) / 60000d;
        }
        set
        {
            _element?.SetAttributeValue("ky", (int)Math.Round(value * 60000d));
        }
    }

    /// <inheritdoc/>
    public bool RotateShadowWithShape
    {
        get
        {
            var val = _element?.Attribute("rotWithShape")?.Value;
            return val is null || val == "1";
        }
        set
        {
            _element?.SetAttributeValue("rotWithShape", value ? "1" : "0");
        }
    }

    /// <inheritdoc/>
    public double ScaleHorizontal
    {
        get
        {
            var val = _element?.Attribute("sx")?.Value;
            return val is null ? 100d : int.Parse(val) / 1000d;
        }
        set
        {
            _element?.SetAttributeValue("sx", (int)Math.Round(value * 1000d));
        }
    }

    /// <inheritdoc/>
    public double ScaleVertical
    {
        get
        {
            var val = _element?.Attribute("sy")?.Value;
            return val is null ? 100d : int.Parse(val) / 1000d;
        }
        set
        {
            _element?.SetAttributeValue("sy", (int)Math.Round(value * 1000d));
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
