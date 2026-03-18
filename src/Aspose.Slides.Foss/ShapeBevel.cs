using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the bevel (relief) properties of a shape's face.
/// </summary>
public sealed class ShapeBevel : PVIObject, IShapeBevel
{
    private const float EmuPerPoint = 12700f;

    private static readonly Dictionary<string, BevelPresetType> BevelMap = new()
    {
        ["angle"] = BevelPresetType.Angle,
        ["artDeco"] = BevelPresetType.ArtDeco,
        ["circle"] = BevelPresetType.Circle,
        ["convex"] = BevelPresetType.Convex,
        ["coolSlant"] = BevelPresetType.CoolSlant,
        ["cross"] = BevelPresetType.Cross,
        ["divot"] = BevelPresetType.Divot,
        ["hardEdge"] = BevelPresetType.HardEdge,
        ["relaxedInset"] = BevelPresetType.RelaxedInset,
        ["riblet"] = BevelPresetType.Riblet,
        ["slope"] = BevelPresetType.Slope,
        ["softRound"] = BevelPresetType.SoftRound,
    };

    private static readonly Dictionary<BevelPresetType, string> BevelMapRev =
        BevelMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private readonly bool _isTop;
    private XElement? _bevelElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes a new instance of the <see cref="ShapeBevel"/> class.
    /// </summary>
    /// <param name="isTopBevel">Whether this bevel represents the top face.</param>
    public ShapeBevel(bool isTopBevel = true)
    {
        _isTop = isTopBevel;
    }

    /// <summary>
    /// Binds this object to a live XML element and its owning slide part.
    /// </summary>
    internal void InitInternal(XElement bevelElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _bevelElement = bevelElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <inheritdoc/>
    public float Width
    {
        get
        {
            var val = _bevelElement?.Attribute("w")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            _bevelElement?.SetAttributeValue("w", (int)Math.Round(value * EmuPerPoint));
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get
        {
            var val = _bevelElement?.Attribute("h")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            _bevelElement?.SetAttributeValue("h", (int)Math.Round(value * EmuPerPoint));
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public BevelPresetType BevelType
    {
        get
        {
            var prst = _bevelElement?.Attribute("prst")?.Value;
            if (prst is not null && BevelMap.TryGetValue(prst, out var type))
                return type;
            return BevelPresetType.NotDefined;
        }
        set
        {
            if (_bevelElement is not null)
            {
                if (value == BevelPresetType.NotDefined)
                {
                    _bevelElement.Attribute("prst")?.Remove();
                }
                else if (BevelMapRev.TryGetValue(value, out var xml))
                {
                    _bevelElement.SetAttributeValue("prst", xml);
                }
            }

            _slidePart?.Save();
        }
    }
}
