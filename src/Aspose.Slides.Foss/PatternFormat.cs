using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a pattern fill format.
/// </summary>
public sealed class PatternFormat : PVIObject, IPatternFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly Dictionary<PatternStyle, string> EnumToXml = new()
    {
        [PatternStyle.Percent05] = "pct5",
        [PatternStyle.Percent10] = "pct10",
        [PatternStyle.Percent20] = "pct20",
        [PatternStyle.Percent25] = "pct25",
        [PatternStyle.Percent30] = "pct30",
        [PatternStyle.Percent40] = "pct40",
        [PatternStyle.Percent50] = "pct50",
        [PatternStyle.Percent60] = "pct60",
        [PatternStyle.Percent70] = "pct70",
        [PatternStyle.Percent75] = "pct75",
        [PatternStyle.Percent80] = "pct80",
        [PatternStyle.Percent90] = "pct90",
        [PatternStyle.DarkHorizontal] = "dkHorz",
        [PatternStyle.DarkVertical] = "dkVert",
        [PatternStyle.DarkDownwardDiagonal] = "dkDnDiag",
        [PatternStyle.DarkUpwardDiagonal] = "dkUpDiag",
        [PatternStyle.SmallCheckerBoard] = "smCheck",
        [PatternStyle.Trellis] = "trellis",
        [PatternStyle.LightHorizontal] = "ltHorz",
        [PatternStyle.LightVertical] = "ltVert",
        [PatternStyle.LightDownwardDiagonal] = "ltDnDiag",
        [PatternStyle.LightUpwardDiagonal] = "ltUpDiag",
        [PatternStyle.SmallGrid] = "smGrid",
        [PatternStyle.DottedDiamond] = "dottedDmnd",
        [PatternStyle.DashedDownwardDiagonal] = "dashDnDiag",
        [PatternStyle.DashedUpwardDiagonal] = "dashUpDiag",
        [PatternStyle.DashedHorizontal] = "dashHorz",
        [PatternStyle.DashedVertical] = "dashVert",
        [PatternStyle.NarrowVertical] = "narVert",
        [PatternStyle.NarrowHorizontal] = "narHorz",
        [PatternStyle.LargeConfetti] = "lgConfetti",
        [PatternStyle.LargeGrid] = "lgGrid",
        [PatternStyle.HorizontalBrick] = "horzBrick",
        [PatternStyle.LargeCheckerBoard] = "lgCheck",
        [PatternStyle.SmallConfetti] = "smConfetti",
        [PatternStyle.Zigzag] = "zigZag",
        [PatternStyle.SolidDiamond] = "solidDmnd",
        [PatternStyle.DiagonalBrick] = "diagBrick",
        [PatternStyle.OutlinedDiamond] = "openDmnd",
        [PatternStyle.Plaid] = "plaid",
        [PatternStyle.Sphere] = "sphere",
        [PatternStyle.Weave] = "weave",
        [PatternStyle.DottedGrid] = "dottedGrid",
        [PatternStyle.Divot] = "divot",
        [PatternStyle.Shingle] = "shingle",
        [PatternStyle.Wave] = "wave",
        [PatternStyle.Horizontal] = "horz",
        [PatternStyle.Vertical] = "vert",
        [PatternStyle.Cross] = "cross",
        [PatternStyle.DownwardDiagonal] = "dnDiag",
        [PatternStyle.UpwardDiagonal] = "upDiag",
        [PatternStyle.DiagonalCross] = "diagCross",
    };

    private static readonly Dictionary<string, PatternStyle> XmlToEnum =
        EnumToXml.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _pattFillElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state.
    /// </summary>
    internal void InitInternal(XElement pattFillElement, IBaseSlide? parentSlide, SlidePart? slidePart = null)
    {
        _pattFillElement = pattFillElement;
        _parentSlide = parentSlide;
        _slidePart = slidePart;
    }

    private void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public PatternStyle PatternStyle
    {
        get
        {
            var val = _pattFillElement?.Attribute("prst")?.Value;
            if (val is null)
                return PatternStyle.NotDefined;
            if (XmlToEnum.TryGetValue(val, out var style))
                return style;
            return PatternStyle.Unknown;
        }
        set
        {
            if (_pattFillElement is null) return;
            if (value is PatternStyle.NotDefined or PatternStyle.Unknown)
            {
                _pattFillElement.Attribute("prst")?.Remove();
            }
            else if (EnumToXml.TryGetValue(value, out var xmlVal))
            {
                _pattFillElement.SetAttributeValue("prst", xmlVal);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IColorFormat ForeColor
    {
        get
        {
            var fgClr = _pattFillElement?.Element(ANs + "fgClr");
            if (fgClr is null && _pattFillElement is not null)
            {
                fgClr = new XElement(ANs + "fgClr");
                _pattFillElement.AddFirst(fgClr);
            }
            var cf = new ColorFormat();
            if (fgClr is not null)
                cf.InitInternal(fgClr, _parentSlide, _slidePart);
            return cf;
        }
    }

    /// <inheritdoc/>
    public IColorFormat BackColor
    {
        get
        {
            var bgClr = _pattFillElement?.Element(ANs + "bgClr");
            if (bgClr is null && _pattFillElement is not null)
            {
                bgClr = new XElement(ANs + "bgClr");
                // bgClr goes after fgClr
                var fgClr = _pattFillElement.Element(ANs + "fgClr");
                if (fgClr is not null)
                    fgClr.AddAfterSelf(bgClr);
                else
                    _pattFillElement.AddFirst(bgClr);
            }
            var cf = new ColorFormat();
            if (bgClr is not null)
                cf.InitInternal(bgClr, _parentSlide, _slidePart);
            return cf;
        }
    }
}
