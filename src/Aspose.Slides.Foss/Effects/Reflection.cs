using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a reflection effect.
/// </summary>
public sealed class Reflection : IReflection
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

    private static readonly Dictionary<RectangleAlignment, string> AlgnMapRev = AlgnMap
        .ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="Reflection"/> class.
    /// </summary>
    public Reflection()
    {
    }

    internal Reflection(XElement element)
    {
        _element = element;
    }

    /// <summary>
    /// Initializes the reflection effect with its XML element, slide part, and parent slide.
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
    public float StartPosAlpha
    {
        get
        {
            var val = Element?.Attribute("stPos")?.Value;
            return val is null ? 0f : int.Parse(val) / 1000f;
        }
        set
        {
            Element?.SetAttributeValue("stPos", (int)MathF.Round(value * 1000f));
        }
    }

    /// <inheritdoc/>
    public float EndPosAlpha
    {
        get
        {
            var val = Element?.Attribute("endPos")?.Value;
            return val is null ? 100f : int.Parse(val) / 1000f;
        }
        set
        {
            Element?.SetAttributeValue("endPos", (int)MathF.Round(value * 1000f));
        }
    }

    /// <inheritdoc/>
    public float FadeDirection
    {
        get
        {
            var val = Element?.Attribute("fadeDir")?.Value;
            return val is null ? 90f : int.Parse(val) / 60000f;
        }
        set
        {
            Element?.SetAttributeValue("fadeDir", (int)MathF.Round(value * 60000f));
        }
    }

    /// <inheritdoc/>
    public float StartReflectionOpacity
    {
        get
        {
            var val = Element?.Attribute("stA")?.Value;
            return val is null ? 100f : int.Parse(val) / 1000f;
        }
        set
        {
            Element?.SetAttributeValue("stA", (int)MathF.Round(value * 1000f));
        }
    }

    /// <inheritdoc/>
    public float EndReflectionOpacity
    {
        get
        {
            var val = Element?.Attribute("endA")?.Value;
            return val is null ? 0f : int.Parse(val) / 1000f;
        }
        set
        {
            Element?.SetAttributeValue("endA", (int)MathF.Round(value * 1000f));
        }
    }

    /// <inheritdoc/>
    public double BlurRadius
    {
        get
        {
            var val = Element?.Attribute("blurRad")?.Value;
            return val is null ? 0d : int.Parse(val) / (double)EmuPerPoint;
        }
        set
        {
            Element?.SetAttributeValue("blurRad", (int)Math.Round(value * EmuPerPoint));
        }
    }

    /// <inheritdoc/>
    public float Direction
    {
        get
        {
            var val = Element?.Attribute("dir")?.Value;
            return val is null ? 0f : int.Parse(val) / 60000f;
        }
        set
        {
            Element?.SetAttributeValue("dir", (int)MathF.Round(value * 60000f));
        }
    }

    /// <inheritdoc/>
    public double Distance
    {
        get
        {
            var val = Element?.Attribute("dist")?.Value;
            return val is null ? 0d : int.Parse(val) / (double)EmuPerPoint;
        }
        set
        {
            Element?.SetAttributeValue("dist", (int)Math.Round(value * EmuPerPoint));
        }
    }

    /// <inheritdoc/>
    public RectangleAlignment RectangleAlign
    {
        get
        {
            var val = Element?.Attribute("algn")?.Value;
            if (val is null)
                return RectangleAlignment.Bottom;
            return AlgnMap.TryGetValue(val, out var result) ? result : RectangleAlignment.NotDefined;
        }
        set
        {
            if (value == RectangleAlignment.NotDefined)
            {
                Element?.Attribute("algn")?.Remove();
            }
            else if (AlgnMapRev.TryGetValue(value, out var ooxmlVal))
            {
                Element?.SetAttributeValue("algn", ooxmlVal);
            }
        }
    }

    /// <inheritdoc/>
    public double SkewHorizontal
    {
        get
        {
            var val = Element?.Attribute("kx")?.Value;
            return val is null ? 0d : int.Parse(val) / 60000d;
        }
        set
        {
            Element?.SetAttributeValue("kx", (int)Math.Round(value * 60000d));
        }
    }

    /// <inheritdoc/>
    public double SkewVertical
    {
        get
        {
            var val = Element?.Attribute("ky")?.Value;
            return val is null ? 0d : int.Parse(val) / 60000d;
        }
        set
        {
            Element?.SetAttributeValue("ky", (int)Math.Round(value * 60000d));
        }
    }

    /// <inheritdoc/>
    public bool RotateShadowWithShape
    {
        get
        {
            var val = Element?.Attribute("rotWithShape")?.Value;
            return val is null || val == "1";
        }
        set
        {
            Element?.SetAttributeValue("rotWithShape", value ? "1" : "0");
        }
    }

    /// <inheritdoc/>
    public double ScaleHorizontal
    {
        get
        {
            var val = Element?.Attribute("sx")?.Value;
            return val is null ? 0d : int.Parse(val) / 1000d;
        }
        set
        {
            Element?.SetAttributeValue("sx", (int)Math.Round(value * 1000d));
        }
    }

    /// <inheritdoc/>
    public double ScaleVertical
    {
        get
        {
            var val = Element?.Attribute("sy")?.Value;
            return val is null ? 0d : int.Parse(val) / 1000d;
        }
        set
        {
            Element?.SetAttributeValue("sy", (int)Math.Round(value * 1000d));
        }
    }

    /// <inheritdoc/>
    public IImageTransformOperation AsIImageTransformOperation => this;
}
