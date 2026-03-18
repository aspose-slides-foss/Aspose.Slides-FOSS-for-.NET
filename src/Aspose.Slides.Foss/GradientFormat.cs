using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a gradient format.
/// </summary>
public sealed class GradientFormat : IGradientFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const float AngleFactor = 60000f;

    private static readonly Dictionary<int, GradientDirection> AngleToDirection = new()
    {
        [0] = GradientDirection.FromCorner1,
        [5400000] = GradientDirection.FromCorner2,
        [10800000] = GradientDirection.FromCorner4,
        [16200000] = GradientDirection.FromCorner3,
    };

    private static readonly Dictionary<GradientDirection, int> DirectionToAngle = AngleToDirection
        .ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _gradFillElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes internal state.
    /// </summary>
    internal void InitInternal(XElement gradFillElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _gradFillElement = gradFillElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <inheritdoc/>
    public TileFlip TileFlip
    {
        get
        {
            var val = _gradFillElement?.Attribute("flip")?.Value;
            return val switch
            {
                "x" => TileFlip.FlipX,
                "y" => TileFlip.FlipY,
                "xy" => TileFlip.FlipBoth,
                "none" => TileFlip.NoFlip,
                _ => TileFlip.NotDefined,
            };
        }
        set
        {
            if (_gradFillElement is null) return;
            var xml = value switch
            {
                TileFlip.FlipX => "x",
                TileFlip.FlipY => "y",
                TileFlip.FlipBoth => "xy",
                TileFlip.NoFlip => "none",
                _ => null,
            };
            if (xml is not null)
                _gradFillElement.SetAttributeValue("flip", xml);
            else
                _gradFillElement.Attribute("flip")?.Remove();
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public GradientDirection GradientDirection
    {
        get
        {
            var lin = _gradFillElement?.Element(ANs + "lin");
            if (lin is not null)
            {
                var angStr = lin.Attribute("ang")?.Value ?? "0";
                if (int.TryParse(angStr, out var ang) && AngleToDirection.TryGetValue(ang, out var dir))
                    return dir;
            }

            var path = _gradFillElement?.Element(ANs + "path");
            if (path is not null)
                return GradientDirection.FromCenter;

            return GradientDirection.NotDefined;
        }
        set
        {
            if (_gradFillElement is null || value == GradientDirection.NotDefined) return;

            if (value == GradientDirection.FromCenter)
            {
                _gradFillElement.Element(ANs + "lin")?.Remove();
                if (_gradFillElement.Element(ANs + "path") is null)
                    _gradFillElement.Add(new XElement(ANs + "path", new XAttribute("path", "circle")));
            }
            else
            {
                _gradFillElement.Element(ANs + "path")?.Remove();
                var ang = DirectionToAngle.GetValueOrDefault(value, 0);
                var lin = _gradFillElement.Element(ANs + "lin");
                if (lin is null)
                {
                    lin = new XElement(ANs + "lin");
                    _gradFillElement.Add(lin);
                }
                lin.SetAttributeValue("ang", ang.ToString());
            }

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public GradientShape GradientShape
    {
        get
        {
            if (_gradFillElement?.Element(ANs + "lin") is not null)
                return GradientShape.Linear;
            var path = _gradFillElement?.Element(ANs + "path");
            if (path is null) return GradientShape.NotDefined;
            return path.Attribute("path")?.Value switch
            {
                "rect" => GradientShape.Rectangle,
                "circle" => GradientShape.Radial,
                "shape" => GradientShape.Path,
                _ => GradientShape.NotDefined,
            };
        }
        set
        {
            if (_gradFillElement is null || value == GradientShape.NotDefined) return;

            _gradFillElement.Element(ANs + "lin")?.Remove();
            _gradFillElement.Element(ANs + "path")?.Remove();

            if (value == GradientShape.Linear)
            {
                _gradFillElement.Add(new XElement(ANs + "lin",
                    new XAttribute("ang", "0"),
                    new XAttribute("scaled", "1")));
            }
            else
            {
                var pathVal = value switch
                {
                    GradientShape.Rectangle => "rect",
                    GradientShape.Radial => "circle",
                    GradientShape.Path => "shape",
                    _ => "rect",
                };
                _gradFillElement.Add(new XElement(ANs + "path", new XAttribute("path", pathVal)));
            }

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public IGradientStopCollection GradientStops
    {
        get
        {
            var gsLst = _gradFillElement?.Element(ANs + "gsLst");
            if (gsLst is null && _gradFillElement is not null)
            {
                gsLst = new XElement(ANs + "gsLst");
                _gradFillElement.AddFirst(gsLst);
            }
            var coll = new GradientStopCollection();
            if (gsLst is not null)
                coll.InitInternal(gsLst, _slidePart, _parentSlide);
            return coll;
        }
    }

    /// <inheritdoc/>
    public float LinearGradientAngle
    {
        get
        {
            var lin = _gradFillElement?.Element(ANs + "lin");
            var ang = lin?.Attribute("ang")?.Value;
            return ang is not null && int.TryParse(ang, out var v) ? v / AngleFactor : 0f;
        }
        set
        {
            if (_gradFillElement is null) return;
            var lin = _gradFillElement.Element(ANs + "lin");
            if (lin is null)
            {
                lin = new XElement(ANs + "lin");
                _gradFillElement.Add(lin);
            }
            lin.SetAttributeValue("ang", (int)MathF.Round(value * AngleFactor));
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public NullableBool LinearGradientScaled
    {
        get
        {
            var lin = _gradFillElement?.Element(ANs + "lin");
            var scaled = lin?.Attribute("scaled")?.Value;
            return scaled switch
            {
                "1" or "true" => NullableBool.True,
                "0" or "false" => NullableBool.False,
                _ => NullableBool.NotDefined,
            };
        }
        set
        {
            if (_gradFillElement is null) return;
            var lin = _gradFillElement.Element(ANs + "lin");
            if (lin is null)
            {
                lin = new XElement(ANs + "lin");
                _gradFillElement.Add(lin);
            }
            if (value == NullableBool.True)
                lin.SetAttributeValue("scaled", "1");
            else if (value == NullableBool.False)
                lin.SetAttributeValue("scaled", "0");
            else
                lin.Attribute("scaled")?.Remove();
            _slidePart?.Save();
        }
    }
}
