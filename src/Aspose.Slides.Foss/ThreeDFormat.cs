using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents 3-D formatting properties for a shape.
/// </summary>
public sealed class ThreeDFormat : PVIObject, IThreeDFormat, IThreeDParamSource
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    private static readonly Dictionary<string, MaterialPresetType> MaterialMap = new()
    {
        ["clear"] = MaterialPresetType.Clear,
        ["dkEdge"] = MaterialPresetType.DkEdge,
        ["flat"] = MaterialPresetType.Flat,
        ["legacyMatte"] = MaterialPresetType.LegacyMatte,
        ["legacyMetal"] = MaterialPresetType.LegacyMetal,
        ["legacyPlastic"] = MaterialPresetType.LegacyPlastic,
        ["legacyWireframe"] = MaterialPresetType.LegacyWireframe,
        ["matte"] = MaterialPresetType.Matte,
        ["metal"] = MaterialPresetType.Metal,
        ["plastic"] = MaterialPresetType.Plastic,
        ["powder"] = MaterialPresetType.Powder,
        ["softEdge"] = MaterialPresetType.SoftEdge,
        ["softmetal"] = MaterialPresetType.Softmetal,
        ["translucentPowder"] = MaterialPresetType.TranslucentPowder,
        ["warmMatte"] = MaterialPresetType.WarmMatte,
    };

    private static readonly Dictionary<MaterialPresetType, string> MaterialMapRev =
        MaterialMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private XElement? _parentElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state for this 3-D format.
    /// </summary>
    internal void InitInternal(XElement parentElement, IBaseSlide? parentSlide)
    {
        _parentElement = parentElement;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Initializes internal state with a slide part for persistence.
    /// </summary>
    internal void InitInternal(XElement parentElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _parentElement = parentElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private XElement? GetSp3d() => _parentElement?.Element(ANs + "sp3d");

    private XElement EnsureSp3d()
    {
        var sp3d = GetSp3d();
        if (sp3d is not null) return sp3d;

        var el = new XElement(ANs + "sp3d");
        if (_parentElement is not null)
        {
            XElement? insertBefore = null;
            foreach (var child in _parentElement.Elements())
            {
                if (child.Name.LocalName == "extLst")
                {
                    insertBefore = child;
                    break;
                }
            }

            if (insertBefore is not null)
                insertBefore.AddBeforeSelf(el);
            else
                _parentElement.Add(el);
        }
        return el;
    }

    private XElement? GetScene3d() => _parentElement?.Element(ANs + "scene3d");

    private XElement EnsureScene3d()
    {
        var scene = GetScene3d();
        if (scene is not null) return scene;

        var el = new XElement(ANs + "scene3d");
        if (_parentElement is not null)
        {
            XElement? insertBefore = null;
            foreach (var child in _parentElement.Elements())
            {
                if (child.Name.LocalName is "sp3d" or "extLst")
                {
                    insertBefore = child;
                    break;
                }
            }

            if (insertBefore is not null)
                insertBefore.AddBeforeSelf(el);
            else
                _parentElement.Add(el);
        }
        return el;
    }

    private void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public float ContourWidth
    {
        get
        {
            var sp3d = GetSp3d();
            if (sp3d is null) return 0f;
            var val = sp3d.Attribute("contourW")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            var sp3d = EnsureSp3d();
            sp3d.SetAttributeValue("contourW", (int)Math.Round(value * EmuPerPoint));
            Save();
        }
    }

    /// <inheritdoc/>
    public float ExtrusionHeight
    {
        get
        {
            var sp3d = GetSp3d();
            if (sp3d is null) return 0f;
            var val = sp3d.Attribute("extrusionH")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            var sp3d = EnsureSp3d();
            sp3d.SetAttributeValue("extrusionH", (int)Math.Round(value * EmuPerPoint));
            Save();
        }
    }

    /// <inheritdoc/>
    public float Depth
    {
        get
        {
            var sp3d = GetSp3d();
            if (sp3d is null) return 0f;
            var val = sp3d.Attribute("z")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            var sp3d = EnsureSp3d();
            sp3d.SetAttributeValue("z", (int)Math.Round(value * EmuPerPoint));
            Save();
        }
    }

    /// <inheritdoc/>
    public IShapeBevel BevelTop
    {
        get
        {
            var sp3d = EnsureSp3d();
            var bevelT = sp3d.Element(ANs + "bevelT");
            if (bevelT is null)
            {
                bevelT = new XElement(ANs + "bevelT");
                sp3d.Add(bevelT);
            }
            var bevel = new ShapeBevel(true);
            bevel.InitInternal(bevelT, _slidePart, _parentSlide);
            return bevel;
        }
    }

    /// <inheritdoc/>
    public IShapeBevel BevelBottom
    {
        get
        {
            var sp3d = EnsureSp3d();
            var bevelB = sp3d.Element(ANs + "bevelB");
            if (bevelB is null)
            {
                bevelB = new XElement(ANs + "bevelB");
                sp3d.Add(bevelB);
            }
            var bevel = new ShapeBevel(false);
            bevel.InitInternal(bevelB, _slidePart, _parentSlide);
            return bevel;
        }
    }

    /// <inheritdoc/>
    public IColorFormat ContourColor
    {
        get
        {
            var sp3d = EnsureSp3d();
            var contourClr = sp3d.Element(ANs + "contourClr");
            if (contourClr is null)
            {
                contourClr = new XElement(ANs + "contourClr");
                sp3d.Add(contourClr);
            }
            var cf = new ColorFormat();
            cf.InitInternal(contourClr, _parentSlide, _slidePart);
            return cf;
        }
    }

    /// <inheritdoc/>
    public IColorFormat ExtrusionColor
    {
        get
        {
            var sp3d = EnsureSp3d();
            var extrusionClr = sp3d.Element(ANs + "extrusionClr");
            if (extrusionClr is null)
            {
                extrusionClr = new XElement(ANs + "extrusionClr");
                sp3d.Add(extrusionClr);
            }
            var cf = new ColorFormat();
            cf.InitInternal(extrusionClr, _parentSlide, _slidePart);
            return cf;
        }
    }

    /// <inheritdoc/>
    public ICamera Camera
    {
        get
        {
            var scene3d = EnsureScene3d();
            var cam = new Camera();
            cam.InitInternal(scene3d, _slidePart, _parentSlide);
            return cam;
        }
    }

    /// <inheritdoc/>
    public ILightRig LightRig
    {
        get
        {
            var scene3d = EnsureScene3d();
            var lr = new LightRig();
            lr.InitInternal(scene3d, _slidePart, _parentSlide);
            return lr;
        }
    }

    /// <inheritdoc/>
    public MaterialPresetType Material
    {
        get
        {
            var sp3d = GetSp3d();
            if (sp3d is null) return MaterialPresetType.NotDefined;
            var val = sp3d.Attribute("prstMaterial")?.Value;
            if (val is not null && MaterialMap.TryGetValue(val, out var type))
                return type;
            return MaterialPresetType.NotDefined;
        }
        set
        {
            var sp3d = EnsureSp3d();
            if (value == MaterialPresetType.NotDefined)
                sp3d.Attribute("prstMaterial")?.Remove();
            else if (MaterialMapRev.TryGetValue(value, out var ooxmlVal))
                sp3d.SetAttributeValue("prstMaterial", ooxmlVal);
            Save();
        }
    }
}
