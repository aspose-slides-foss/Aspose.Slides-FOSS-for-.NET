using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss.Effects;

/// <summary>
/// Represents a preset shadow effect.
/// </summary>
public sealed class PresetShadow : IPresetShadow
{
    private const float EmuPerPoint = 12700f;

    private XElement? _element;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    internal XElement? Element => _element;

    /// <summary>
    /// Initializes a new instance of the <see cref="PresetShadow"/> class.
    /// </summary>
    public PresetShadow()
    {
    }

    internal PresetShadow(XElement element)
    {
        _element = element;
    }

    /// <summary>
    /// Initializes the preset shadow effect with its XML element, slide part, and parent slide.
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
    public PresetShadowType Preset
    {
        get
        {
            var val = _element?.Attribute("prst")?.Value;
            if (val is null) return PresetShadowType.TopLeftDropShadow;
            return PrstToEnum.GetValueOrDefault(val, PresetShadowType.TopLeftDropShadow);
        }
        set
        {
            if (EnumToPrst.TryGetValue(value, out var ooxmlVal))
            {
                _element?.SetAttributeValue("prst", ooxmlVal);
            }
        }
    }

    private static readonly Dictionary<string, PresetShadowType> PrstToEnum = new()
    {
        ["shdw1"] = PresetShadowType.TopLeftDropShadow,
        ["shdw2"] = PresetShadowType.TopLeftLargeDropShadow,
        ["shdw3"] = PresetShadowType.BackLeftLongPerspectiveShadow,
        ["shdw4"] = PresetShadowType.BackRightLongPerspectiveShadow,
        ["shdw5"] = PresetShadowType.TopLeftDoubleDropShadow,
        ["shdw6"] = PresetShadowType.BottomRightSmallDropShadow,
        ["shdw7"] = PresetShadowType.FrontLeftLongPerspectiveShadow,
        ["shdw8"] = PresetShadowType.FrontRightLongPerspectiveShadow,
        ["shdw9"] = PresetShadowType.OuterBoxShadow3D,
        ["shdw10"] = PresetShadowType.InnerBoxShadow3D,
        ["shdw11"] = PresetShadowType.BackCenterPerspectiveShadow,
        ["shdw12"] = PresetShadowType.TopRightDropShadow,
        ["shdw13"] = PresetShadowType.FrontBottomShadow,
        ["shdw14"] = PresetShadowType.BackLeftPerspectiveShadow,
        ["shdw15"] = PresetShadowType.BackRightPerspectiveShadow,
        ["shdw16"] = PresetShadowType.BottomLeftDropShadow,
        ["shdw17"] = PresetShadowType.BottomRightDropShadow,
        ["shdw18"] = PresetShadowType.FrontLeftPerspectiveShadow,
        ["shdw19"] = PresetShadowType.FrontRightPerspectiveShadow,
        ["shdw20"] = PresetShadowType.TopLeftSmallDropShadow,
    };

    private static readonly Dictionary<PresetShadowType, string> EnumToPrst =
        PrstToEnum.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    /// <inheritdoc/>
    public IImageTransformOperation AsIImageTransformOperation => this;
}
