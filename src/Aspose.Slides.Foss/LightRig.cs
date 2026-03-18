using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a light rig.
/// </summary>
public sealed class LightRig : PVIObject, ILightRig
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float RotationUnit = 60000f;

    private static readonly Dictionary<string, LightRigPresetType> LightTypeMap = new()
    {
        ["balanced"] = LightRigPresetType.Balanced,
        ["brightRm"] = LightRigPresetType.BrightRoom,
        ["chilly"] = LightRigPresetType.Chilly,
        ["contrasting"] = LightRigPresetType.Contrasting,
        ["flat"] = LightRigPresetType.Flat,
        ["flood"] = LightRigPresetType.Flood,
        ["freezing"] = LightRigPresetType.Freezing,
        ["glow"] = LightRigPresetType.Glow,
        ["harsh"] = LightRigPresetType.Harsh,
        ["legacyFlat1"] = LightRigPresetType.LegacyFlat1,
        ["legacyFlat2"] = LightRigPresetType.LegacyFlat2,
        ["legacyFlat3"] = LightRigPresetType.LegacyFlat3,
        ["legacyFlat4"] = LightRigPresetType.LegacyFlat4,
        ["legacyHarsh1"] = LightRigPresetType.LegacyHarsh1,
        ["legacyHarsh2"] = LightRigPresetType.LegacyHarsh2,
        ["legacyHarsh3"] = LightRigPresetType.LegacyHarsh3,
        ["legacyHarsh4"] = LightRigPresetType.LegacyHarsh4,
        ["legacyNormal1"] = LightRigPresetType.LegacyNormal1,
        ["legacyNormal2"] = LightRigPresetType.LegacyNormal2,
        ["legacyNormal3"] = LightRigPresetType.LegacyNormal3,
        ["legacyNormal4"] = LightRigPresetType.LegacyNormal4,
        ["morning"] = LightRigPresetType.Morning,
        ["soft"] = LightRigPresetType.Soft,
        ["sunrise"] = LightRigPresetType.Sunrise,
        ["sunset"] = LightRigPresetType.Sunset,
        ["threePt"] = LightRigPresetType.ThreePt,
        ["twoPt"] = LightRigPresetType.TwoPt,
    };

    private static readonly Dictionary<LightRigPresetType, string> LightTypeMapRev =
        LightTypeMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LightingDirection> DirMap = new()
    {
        ["tl"] = LightingDirection.TopLeft,
        ["t"] = LightingDirection.Top,
        ["tr"] = LightingDirection.TopRight,
        ["r"] = LightingDirection.Right,
        ["br"] = LightingDirection.BottomRight,
        ["b"] = LightingDirection.Bottom,
        ["bl"] = LightingDirection.BottomLeft,
        ["l"] = LightingDirection.Left,
    };

    private static readonly Dictionary<LightingDirection, string> DirMapRev =
        DirMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private XElement? _scene3d;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state.
    /// </summary>
    internal void InitInternal(XElement scene3dElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _scene3d = scene3dElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private XElement? GetLightRig() => _scene3d?.Element(ANs + "lightRig");

    private XElement EnsureLightRig()
    {
        var lr = GetLightRig();
        if (lr is not null) return lr;
        lr = new XElement(ANs + "lightRig",
            new XAttribute("rig", "threePt"),
            new XAttribute("dir", "t"));
        _scene3d?.Add(lr);
        return lr;
    }

    private void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public LightingDirection Direction
    {
        get
        {
            var lr = GetLightRig();
            if (lr is null) return LightingDirection.NotDefined;
            var val = lr.Attribute("dir")?.Value;
            if (val is not null && DirMap.TryGetValue(val, out var dir))
                return dir;
            return LightingDirection.NotDefined;
        }
        set
        {
            var lr = EnsureLightRig();
            if (value == LightingDirection.NotDefined)
                lr.Attribute("dir")?.Remove();
            else if (DirMapRev.TryGetValue(value, out var ooxmlVal))
                lr.SetAttributeValue("dir", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public LightRigPresetType LightType
    {
        get
        {
            var lr = GetLightRig();
            if (lr is null) return LightRigPresetType.NotDefined;
            var val = lr.Attribute("rig")?.Value;
            if (val is not null && LightTypeMap.TryGetValue(val, out var type))
                return type;
            return LightRigPresetType.NotDefined;
        }
        set
        {
            var lr = EnsureLightRig();
            if (value == LightRigPresetType.NotDefined)
                lr.Attribute("rig")?.Remove();
            else if (LightTypeMapRev.TryGetValue(value, out var ooxmlVal))
                lr.SetAttributeValue("rig", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public void SetRotation(float latitude, float longitude, float revolution)
    {
        var lr = EnsureLightRig();
        var rot = lr.Element(ANs + "rot");
        if (rot is null)
        {
            rot = new XElement(ANs + "rot");
            lr.Add(rot);
        }
        rot.SetAttributeValue("lat", (int)MathF.Round(latitude * RotationUnit));
        rot.SetAttributeValue("lon", (int)MathF.Round(longitude * RotationUnit));
        rot.SetAttributeValue("rev", (int)MathF.Round(revolution * RotationUnit));
        Save();
    }

    /// <inheritdoc/>
    public float[] GetRotation()
    {
        var lr = GetLightRig();
        var rot = lr?.Element(ANs + "rot");
        if (rot is null) return [0f, 0f, 0f];

        float lat = 0f, lon = 0f, rev = 0f;
        if (int.TryParse(rot.Attribute("lat")?.Value, out var latVal)) lat = latVal / RotationUnit;
        if (int.TryParse(rot.Attribute("lon")?.Value, out var lonVal)) lon = lonVal / RotationUnit;
        if (int.TryParse(rot.Attribute("rev")?.Value, out var revVal)) rev = revVal / RotationUnit;
        return [lat, lon, rev];
    }
}
