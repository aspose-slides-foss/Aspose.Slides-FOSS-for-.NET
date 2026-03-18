using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents 3D camera properties for a shape.
/// </summary>
public sealed class Camera : PVIObject, ICamera
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const float AngleFactor = 60000f;

    private static readonly Dictionary<string, CameraPresetType> CameraPrstMap = new()
    {
        ["isometricBottomDown"] = CameraPresetType.IsometricBottomDown,
        ["isometricBottomUp"] = CameraPresetType.IsometricBottomUp,
        ["isometricLeftDown"] = CameraPresetType.IsometricLeftDown,
        ["isometricLeftUp"] = CameraPresetType.IsometricLeftUp,
        ["isometricOffAxis1Left"] = CameraPresetType.IsometricOffAxis1Left,
        ["isometricOffAxis1Right"] = CameraPresetType.IsometricOffAxis1Right,
        ["isometricOffAxis1Top"] = CameraPresetType.IsometricOffAxis1Top,
        ["isometricOffAxis2Left"] = CameraPresetType.IsometricOffAxis2Left,
        ["isometricOffAxis2Right"] = CameraPresetType.IsometricOffAxis2Right,
        ["isometricOffAxis2Top"] = CameraPresetType.IsometricOffAxis2Top,
        ["isometricOffAxis3Bottom"] = CameraPresetType.IsometricOffAxis3Bottom,
        ["isometricOffAxis3Left"] = CameraPresetType.IsometricOffAxis3Left,
        ["isometricOffAxis3Right"] = CameraPresetType.IsometricOffAxis3Right,
        ["isometricOffAxis4Bottom"] = CameraPresetType.IsometricOffAxis4Bottom,
        ["isometricOffAxis4Left"] = CameraPresetType.IsometricOffAxis4Left,
        ["isometricOffAxis4Right"] = CameraPresetType.IsometricOffAxis4Right,
        ["isometricRightDown"] = CameraPresetType.IsometricRightDown,
        ["isometricRightUp"] = CameraPresetType.IsometricRightUp,
        ["isometricTopDown"] = CameraPresetType.IsometricTopDown,
        ["isometricTopUp"] = CameraPresetType.IsometricTopUp,
        ["legacyObliqueBottom"] = CameraPresetType.LegacyObliqueBottom,
        ["legacyObliqueBottomLeft"] = CameraPresetType.LegacyObliqueBottomLeft,
        ["legacyObliqueBottomRight"] = CameraPresetType.LegacyObliqueBottomRight,
        ["legacyObliqueFront"] = CameraPresetType.LegacyObliqueFront,
        ["legacyObliqueLeft"] = CameraPresetType.LegacyObliqueLeft,
        ["legacyObliqueRight"] = CameraPresetType.LegacyObliqueRight,
        ["legacyObliqueTop"] = CameraPresetType.LegacyObliqueTop,
        ["legacyObliqueTopLeft"] = CameraPresetType.LegacyObliqueTopLeft,
        ["legacyObliqueTopRight"] = CameraPresetType.LegacyObliqueTopRight,
        ["legacyPerspectiveBottom"] = CameraPresetType.LegacyPerspectiveBottom,
        ["legacyPerspectiveBottomLeft"] = CameraPresetType.LegacyPerspectiveBottomLeft,
        ["legacyPerspectiveBottomRight"] = CameraPresetType.LegacyPerspectiveBottomRight,
        ["legacyPerspectiveFront"] = CameraPresetType.LegacyPerspectiveFront,
        ["legacyPerspectiveLeft"] = CameraPresetType.LegacyPerspectiveLeft,
        ["legacyPerspectiveRight"] = CameraPresetType.LegacyPerspectiveRight,
        ["legacyPerspectiveTop"] = CameraPresetType.LegacyPerspectiveTop,
        ["legacyPerspectiveTopLeft"] = CameraPresetType.LegacyPerspectiveTopLeft,
        ["legacyPerspectiveTopRight"] = CameraPresetType.LegacyPerspectiveTopRight,
        ["obliqueBottom"] = CameraPresetType.ObliqueBottom,
        ["obliqueBottomLeft"] = CameraPresetType.ObliqueBottomLeft,
        ["obliqueBottomRight"] = CameraPresetType.ObliqueBottomRight,
        ["obliqueLeft"] = CameraPresetType.ObliqueLeft,
        ["obliqueRight"] = CameraPresetType.ObliqueRight,
        ["obliqueTop"] = CameraPresetType.ObliqueTop,
        ["obliqueTopLeft"] = CameraPresetType.ObliqueTopLeft,
        ["obliqueTopRight"] = CameraPresetType.ObliqueTopRight,
        ["orthographicFront"] = CameraPresetType.OrthographicFront,
        ["perspectiveAbove"] = CameraPresetType.PerspectiveAbove,
        ["perspectiveAboveLeftFacing"] = CameraPresetType.PerspectiveAboveLeftFacing,
        ["perspectiveAboveRightFacing"] = CameraPresetType.PerspectiveAboveRightFacing,
        ["perspectiveBelow"] = CameraPresetType.PerspectiveBelow,
        ["perspectiveContrastingLeftFacing"] = CameraPresetType.PerspectiveContrastingLeftFacing,
        ["perspectiveContrastingRightFacing"] = CameraPresetType.PerspectiveContrastingRightFacing,
        ["perspectiveFront"] = CameraPresetType.PerspectiveFront,
        ["perspectiveHeroicExtremeLeftFacing"] = CameraPresetType.PerspectiveHeroicExtremeLeftFacing,
        ["perspectiveHeroicExtremeRightFacing"] = CameraPresetType.PerspectiveHeroicExtremeRightFacing,
        ["perspectiveHeroicLeftFacing"] = CameraPresetType.PerspectiveHeroicLeftFacing,
        ["perspectiveHeroicRightFacing"] = CameraPresetType.PerspectiveHeroicRightFacing,
        ["perspectiveLeft"] = CameraPresetType.PerspectiveLeft,
        ["perspectiveRelaxed"] = CameraPresetType.PerspectiveRelaxed,
        ["perspectiveRelaxedModerately"] = CameraPresetType.PerspectiveRelaxedModerately,
        ["perspectiveRight"] = CameraPresetType.PerspectiveRight,
    };

    private static readonly Dictionary<CameraPresetType, string> ReverseCameraPrstMap =
        CameraPrstMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private XElement? _scene3d;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state from a scene3d element.
    /// </summary>
    internal void InitInternal(XElement scene3dElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _scene3d = scene3dElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Gets the camera element from the scene3d element, or <c>null</c> if none exists.
    /// </summary>
    internal XElement? GetCamera() => _scene3d?.Element(ANs + "camera");

    /// <summary>
    /// Gets the camera element, creating a default one if it does not exist.
    /// </summary>
    internal XElement EnsureCamera()
    {
        var cam = GetCamera();
        if (cam is not null) return cam;

        var el = new XElement(ANs + "camera", new XAttribute("prst", "orthographicFront"));
        _scene3d?.Add(el);
        return el;
    }

    /// <summary>
    /// Persists changes to the underlying slide part.
    /// </summary>
    internal void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public CameraPresetType CameraType
    {
        get
        {
            var cam = GetCamera();
            var prst = cam?.Attribute("prst")?.Value;
            if (prst is not null && CameraPrstMap.TryGetValue(prst, out var type))
                return type;
            return CameraPresetType.NotDefined;
        }
        set
        {
            var cam = EnsureCamera();
            if (value == CameraPresetType.NotDefined)
            {
                cam.Attribute("prst")?.Remove();
            }
            else if (ReverseCameraPrstMap.TryGetValue(value, out var xml))
            {
                cam.SetAttributeValue("prst", xml);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public float FieldOfViewAngle
    {
        get
        {
            var cam = GetCamera();
            var fov = cam?.Attribute("fov")?.Value;
            return fov is not null && int.TryParse(fov, out var v) ? v / AngleFactor : 0f;
        }
        set
        {
            var cam = EnsureCamera();
            cam.SetAttributeValue("fov", (int)Math.Round(value * AngleFactor));
            Save();
        }
    }

    /// <inheritdoc/>
    public float Zoom
    {
        get
        {
            var cam = GetCamera();
            var zoom = cam?.Attribute("zoom")?.Value;
            return zoom is not null && int.TryParse(zoom, out var v) ? v / 1000f : 100f;
        }
        set
        {
            var cam = EnsureCamera();
            cam.SetAttributeValue("zoom", (int)Math.Round(value * 1000));
            Save();
        }
    }

    /// <inheritdoc/>
    public void SetRotation(float latitude, float longitude, float revolution)
    {
        var cam = EnsureCamera();

        var rot = cam.Element(ANs + "rot");
        if (rot is null)
        {
            rot = new XElement(ANs + "rot");
            cam.Add(rot);
        }

        rot.SetAttributeValue("lat", (int)Math.Round(latitude * AngleFactor));
        rot.SetAttributeValue("lon", (int)Math.Round(longitude * AngleFactor));
        rot.SetAttributeValue("rev", (int)Math.Round(revolution * AngleFactor));
        Save();
    }

    /// <inheritdoc/>
    public float[] GetRotation()
    {
        var cam = GetCamera();
        var rot = cam?.Element(ANs + "rot");
        if (rot is null) return [0f, 0f, 0f];

        float lat = 0f, lon = 0f, rev = 0f;
        if (int.TryParse(rot.Attribute("lat")?.Value, out var latVal)) lat = latVal / AngleFactor;
        if (int.TryParse(rot.Attribute("lon")?.Value, out var lonVal)) lon = lonVal / AngleFactor;
        if (int.TryParse(rot.Attribute("rev")?.Value, out var revVal)) rev = revVal / AngleFactor;
        return [lat, lon, rev];
    }
}
