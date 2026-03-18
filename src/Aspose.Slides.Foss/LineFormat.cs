using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents line formatting properties.
/// </summary>
public sealed class LineFormat : ILineFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    private static readonly Dictionary<string, LineDashStyle> DashMap = new()
    {
        ["solid"] = LineDashStyle.Solid,
        ["dot"] = LineDashStyle.Dot,
        ["dash"] = LineDashStyle.Dash,
        ["lgDash"] = LineDashStyle.LargeDash,
        ["dashDot"] = LineDashStyle.DashDot,
        ["lgDashDot"] = LineDashStyle.LargeDashDot,
        ["lgDashDotDot"] = LineDashStyle.LargeDashDotDot,
        ["sysDash"] = LineDashStyle.SystemDash,
        ["sysDot"] = LineDashStyle.SystemDot,
        ["sysDashDot"] = LineDashStyle.SystemDashDot,
        ["sysDashDotDot"] = LineDashStyle.SystemDashDotDot,
    };

    private static readonly Dictionary<LineDashStyle, string> DashMapRev =
        DashMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineCapStyle> CapMap = new()
    {
        ["rnd"] = LineCapStyle.Round,
        ["sq"] = LineCapStyle.Square,
        ["flat"] = LineCapStyle.Flat,
    };

    private static readonly Dictionary<LineCapStyle, string> CapMapRev =
        CapMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineAlignment> AlgnMap = new()
    {
        ["ctr"] = LineAlignment.Center,
        ["in"] = LineAlignment.Inset,
    };

    private static readonly Dictionary<LineAlignment, string> AlgnMapRev =
        AlgnMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineArrowheadStyle> ArrowTypeMap = new()
    {
        ["none"] = LineArrowheadStyle.None,
        ["triangle"] = LineArrowheadStyle.Triangle,
        ["stealth"] = LineArrowheadStyle.Stealth,
        ["diamond"] = LineArrowheadStyle.Diamond,
        ["oval"] = LineArrowheadStyle.Oval,
        ["arrow"] = LineArrowheadStyle.Open,
    };

    private static readonly Dictionary<LineArrowheadStyle, string> ArrowTypeMapRev =
        ArrowTypeMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineArrowheadWidth> ArrowWMap = new()
    {
        ["sm"] = LineArrowheadWidth.Narrow,
        ["med"] = LineArrowheadWidth.Medium,
        ["lg"] = LineArrowheadWidth.Wide,
    };

    private static readonly Dictionary<LineArrowheadWidth, string> ArrowWMapRev =
        ArrowWMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineArrowheadLength> ArrowLenMap = new()
    {
        ["sm"] = LineArrowheadLength.Short,
        ["med"] = LineArrowheadLength.Medium,
        ["lg"] = LineArrowheadLength.Long,
    };

    private static readonly Dictionary<LineArrowheadLength, string> ArrowLenMapRev =
        ArrowLenMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, LineStyle> CmpdMap = new()
    {
        ["sng"] = LineStyle.Single,
        ["dbl"] = LineStyle.ThinThin,
        ["thickThin"] = LineStyle.ThickThin,
        ["thinThick"] = LineStyle.ThinThick,
        ["tri"] = LineStyle.ThickBetweenThin,
    };

    private static readonly Dictionary<LineStyle, string> CmpdMapRev =
        CmpdMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private XElement? _parentElement;
    private XName? _lnTag;
    private IBaseSlide? _parentSlide;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state for this line format.
    /// </summary>
    internal void InitInternal(XElement parentElement, IBaseSlide? parentSlide, XName? lnTag = null)
    {
        _parentElement = parentElement;
        _parentSlide = parentSlide;
        _lnTag = lnTag;
    }

    /// <summary>
    /// Initializes internal state with a slide part for persistence.
    /// </summary>
    internal void InitInternal(XElement parentElement, SlidePart? slidePart, IBaseSlide? parentSlide, XName? lnTag = null)
    {
        _parentElement = parentElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _lnTag = lnTag;
    }

    private XName LnTag => _lnTag ?? ANs + "ln";

    /// <summary>
    /// Gets the line element if it exists.
    /// </summary>
    internal XElement? GetLn() => _parentElement?.Element(LnTag);

    // OOXML CT_TableCellProperties child order for border lines within <a:tcPr>:
    // lnL  lnR  lnT  lnB  lnTlToBr  lnBlToTr  fill  cell3D  extLst
    private static readonly string[] TcPrChildOrder =
    [
        "lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr",
    ];

    /// <summary>
    /// Gets or creates the line element at the correct OOXML position.
    /// </summary>
    internal XElement EnsureLn()
    {
        var ln = GetLn();
        if (ln is not null) return ln;

        var tag = LnTag;
        var el = new XElement(tag);

        if (_parentElement is null)
            return el;

        // For table cell border elements, use the tcPr child ordering
        var localName = tag.LocalName;
        var tcRank = Array.IndexOf(TcPrChildOrder, localName);
        if (tcRank >= 0)
        {
            foreach (var child in _parentElement.Elements().ToList())
            {
                var childRank = Array.IndexOf(TcPrChildOrder, child.Name.LocalName);
                if (childRank < 0)
                {
                    // Non-border child (fill, cell3D, extLst) comes after all borders
                    child.AddBeforeSelf(el);
                    return el;
                }
                if (childRank > tcRank)
                {
                    child.AddBeforeSelf(el);
                    return el;
                }
            }
            _parentElement.Add(el);
            return el;
        }

        // Default: insert after fill elements, before effects
        XElement? insertBefore = null;
        foreach (var child in _parentElement.Elements())
        {
            var local = child.Name.LocalName;
            if (local is "effectLst" or "effectDag" or "scene3d" or "sp3d" or "extLst")
            {
                insertBefore = child;
                break;
            }
        }

        if (insertBefore is not null)
            insertBefore.AddBeforeSelf(el);
        else
            _parentElement.Add(el);

        return el;
    }

    /// <summary>
    /// Persists changes to the underlying slide part.
    /// </summary>
    internal void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public bool IsFormatNotDefined
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return true;
            return !ln.HasAttributes && !ln.HasElements;
        }
    }

    /// <inheritdoc/>
    public ILineFillFormat FillFormat
    {
        get
        {
            var ln = EnsureLn();
            var lff = new LineFillFormat();
            lff.InitInternal(ln, _slidePart, _parentSlide);
            return lff;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return 0.75f;
            var w = ln.Attribute("w")?.Value;
            if (w is null) return 0.75f;
            return int.TryParse(w, out var v) ? v / EmuPerPoint : 0.75f;
        }
        set
        {
            var ln = EnsureLn();
            ln.SetAttributeValue("w", (int)Math.Round(value * EmuPerPoint));
            Save();
        }
    }

    /// <inheritdoc/>
    public LineDashStyle DashStyle
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return LineDashStyle.NotDefined;
            if (ln.Element(ANs + "custDash") is not null) return LineDashStyle.Custom;
            var prstDash = ln.Element(ANs + "prstDash");
            if (prstDash is null) return LineDashStyle.NotDefined;
            var val = prstDash.Attribute("val")?.Value;
            if (val is not null && DashMap.TryGetValue(val, out var style))
                return style;
            return LineDashStyle.NotDefined;
        }
        set
        {
            var ln = EnsureLn();
            ln.Element(ANs + "prstDash")?.Remove();
            ln.Element(ANs + "custDash")?.Remove();

            if (value == LineDashStyle.NotDefined)
                return;

            if (value == LineDashStyle.Custom)
            {
                InsertLnChild(ln, ANs + "custDash");
            }
            else if (DashMapRev.TryGetValue(value, out var ooxmlVal))
            {
                InsertLnChild(ln, ANs + "prstDash", ("val", ooxmlVal));
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public IList<float> CustomDashPattern
    {
        get
        {
            var ln = GetLn();
            var custDash = ln?.Element(ANs + "custDash");
            if (custDash is null) return [];
            var result = new List<float>();
            foreach (var ds in custDash.Elements(ANs + "ds"))
            {
                var d = ds.Attribute("d")?.Value;
                var sp = ds.Attribute("sp")?.Value;
                if (d is not null && int.TryParse(d, out var dVal))
                    result.Add(dVal / 100000f);
                if (sp is not null && int.TryParse(sp, out var spVal))
                    result.Add(spVal / 100000f);
            }
            return result;
        }
        set
        {
            var ln = EnsureLn();
            ln.Element(ANs + "prstDash")?.Remove();
            ln.Element(ANs + "custDash")?.Remove();

            var custDash = InsertLnChild(ln, ANs + "custDash");
            for (var i = 0; i + 1 < value.Count; i += 2)
            {
                var ds = new XElement(ANs + "ds");
                ds.SetAttributeValue("d", (int)Math.Round(value[i] * 100000f));
                ds.SetAttributeValue("sp", (int)Math.Round(value[i + 1] * 100000f));
                custDash.Add(ds);
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public LineStyle Style
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return LineStyle.NotDefined;
            var val = ln.Attribute("cmpd")?.Value;
            if (val is not null && CmpdMap.TryGetValue(val, out var style))
                return style;
            return LineStyle.NotDefined;
        }
        set
        {
            var ln = EnsureLn();
            if (value == LineStyle.NotDefined)
                ln.Attribute("cmpd")?.Remove();
            else if (CmpdMapRev.TryGetValue(value, out var ooxmlVal))
                ln.SetAttributeValue("cmpd", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public LineCapStyle CapStyle
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return LineCapStyle.NotDefined;
            var val = ln.Attribute("cap")?.Value;
            if (val is not null && CapMap.TryGetValue(val, out var style))
                return style;
            return LineCapStyle.NotDefined;
        }
        set
        {
            var ln = EnsureLn();
            if (value == LineCapStyle.NotDefined)
                ln.Attribute("cap")?.Remove();
            else if (CapMapRev.TryGetValue(value, out var ooxmlVal))
                ln.SetAttributeValue("cap", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public LineAlignment Alignment
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return LineAlignment.NotDefined;
            var val = ln.Attribute("algn")?.Value;
            if (val is not null && AlgnMap.TryGetValue(val, out var align))
                return align;
            return LineAlignment.NotDefined;
        }
        set
        {
            var ln = EnsureLn();
            if (value == LineAlignment.NotDefined)
                ln.Attribute("algn")?.Remove();
            else if (AlgnMapRev.TryGetValue(value, out var ooxmlVal))
                ln.SetAttributeValue("algn", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public LineArrowheadStyle BeginArrowheadStyle
    {
        get => GetArrowStyle(ANs + "headEnd");
        set => SetArrowStyle(ANs + "headEnd", value);
    }

    /// <inheritdoc/>
    public LineArrowheadStyle EndArrowheadStyle
    {
        get => GetArrowStyle(ANs + "tailEnd");
        set => SetArrowStyle(ANs + "tailEnd", value);
    }

    /// <inheritdoc/>
    public LineArrowheadWidth BeginArrowheadWidth
    {
        get => GetArrowWidth(ANs + "headEnd");
        set => SetArrowWidth(ANs + "headEnd", value);
    }

    /// <inheritdoc/>
    public LineArrowheadWidth EndArrowheadWidth
    {
        get => GetArrowWidth(ANs + "tailEnd");
        set => SetArrowWidth(ANs + "tailEnd", value);
    }

    /// <inheritdoc/>
    public LineArrowheadLength BeginArrowheadLength
    {
        get => GetArrowLength(ANs + "headEnd");
        set => SetArrowLength(ANs + "headEnd", value);
    }

    /// <inheritdoc/>
    public LineArrowheadLength EndArrowheadLength
    {
        get => GetArrowLength(ANs + "tailEnd");
        set => SetArrowLength(ANs + "tailEnd", value);
    }

    /// <inheritdoc/>
    public LineJoinStyle JoinStyle
    {
        get
        {
            var ln = GetLn();
            if (ln is null) return LineJoinStyle.NotDefined;
            if (ln.Element(ANs + "round") is not null) return LineJoinStyle.Round;
            if (ln.Element(ANs + "bevel") is not null) return LineJoinStyle.Bevel;
            if (ln.Element(ANs + "miter") is not null) return LineJoinStyle.Miter;
            return LineJoinStyle.NotDefined;
        }
        set
        {
            var ln = EnsureLn();
            ln.Element(ANs + "round")?.Remove();
            ln.Element(ANs + "bevel")?.Remove();
            ln.Element(ANs + "miter")?.Remove();

            if (value == LineJoinStyle.NotDefined)
                return;

            var tagName = value switch
            {
                LineJoinStyle.Round => "round",
                LineJoinStyle.Bevel => "bevel",
                LineJoinStyle.Miter => "miter",
                _ => null,
            };

            if (tagName is not null)
                InsertLnChild(ln, ANs + tagName);

            Save();
        }
    }

    /// <inheritdoc/>
    public float MiterLimit
    {
        get
        {
            var ln = GetLn();
            var miter = ln?.Element(ANs + "miter");
            if (miter is null) return 0;
            var lim = miter.Attribute("lim")?.Value;
            if (lim is not null && int.TryParse(lim, out var v))
                return v / 100000f;
            return 0;
        }
        set
        {
            var ln = EnsureLn();
            // Ensure miter join element exists
            var miter = ln.Element(ANs + "miter");
            if (miter is null)
            {
                ln.Element(ANs + "round")?.Remove();
                ln.Element(ANs + "bevel")?.Remove();
                miter = InsertLnChild(ln, ANs + "miter");
            }
            miter.SetAttributeValue("lim", (int)Math.Round(value * 100000f));
            Save();
        }
    }

    /// <summary>
    /// Gets an arrowhead attribute value from the specified end element.
    /// </summary>
    /// <param name="endTag">The end element tag (e.g., headEnd or tailEnd).</param>
    /// <param name="attr">The attribute name to read.</param>
    /// <returns>The attribute value, or <c>null</c> if not found.</returns>
    internal string? GetArrowAttr(XName endTag, string attr)
    {
        var ln = GetLn();
        if (ln is null) return null;
        var endElem = ln.Element(endTag);
        if (endElem is null) return null;
        return endElem.Attribute(attr)?.Value;
    }

    /// <summary>
    /// Sets an arrowhead attribute value on the specified end element.
    /// </summary>
    /// <param name="endTag">The end element tag (e.g., headEnd or tailEnd).</param>
    /// <param name="attr">The attribute name to set.</param>
    /// <param name="value">The attribute value, or <c>null</c> to remove.</param>
    internal void SetArrowAttr(XName endTag, string attr, string? value)
    {
        var ln = EnsureLn();
        var endElem = ln.Element(endTag);
        if (endElem is null)
            endElem = InsertLnChild(ln, endTag);
        if (value is null)
            endElem.Attribute(attr)?.Remove();
        else
            endElem.SetAttributeValue(attr, value);
        Save();
    }

    private LineArrowheadStyle GetArrowStyle(XName endTag)
    {
        var ln = GetLn();
        var val = ln?.Element(endTag)?.Attribute("type")?.Value;
        if (val is not null && ArrowTypeMap.TryGetValue(val, out var style))
            return style;
        return LineArrowheadStyle.NotDefined;
    }

    private void SetArrowStyle(XName endTag, LineArrowheadStyle value)
    {
        var ln = EnsureLn();
        var endElem = ln.Element(endTag);
        if (endElem is null)
            endElem = InsertLnChild(ln, endTag);
        if (value == LineArrowheadStyle.NotDefined)
            endElem.Attribute("type")?.Remove();
        else if (ArrowTypeMapRev.TryGetValue(value, out var ooxmlVal))
            endElem.SetAttributeValue("type", ooxmlVal);
        Save();
    }

    private LineArrowheadWidth GetArrowWidth(XName endTag)
    {
        var ln = GetLn();
        var val = ln?.Element(endTag)?.Attribute("w")?.Value;
        if (val is not null && ArrowWMap.TryGetValue(val, out var width))
            return width;
        return LineArrowheadWidth.NotDefined;
    }

    private void SetArrowWidth(XName endTag, LineArrowheadWidth value)
    {
        var ln = EnsureLn();
        var endElem = ln.Element(endTag);
        if (endElem is null)
            endElem = InsertLnChild(ln, endTag);
        if (value == LineArrowheadWidth.NotDefined)
            endElem.Attribute("w")?.Remove();
        else if (ArrowWMapRev.TryGetValue(value, out var ooxmlVal))
            endElem.SetAttributeValue("w", ooxmlVal);
        Save();
    }

    private LineArrowheadLength GetArrowLength(XName endTag)
    {
        var ln = GetLn();
        var val = ln?.Element(endTag)?.Attribute("len")?.Value;
        if (val is not null && ArrowLenMap.TryGetValue(val, out var length))
            return length;
        return LineArrowheadLength.NotDefined;
    }

    private void SetArrowLength(XName endTag, LineArrowheadLength value)
    {
        var ln = EnsureLn();
        var endElem = ln.Element(endTag);
        if (endElem is null)
            endElem = InsertLnChild(ln, endTag);
        if (value == LineArrowheadLength.NotDefined)
            endElem.Attribute("len")?.Remove();
        else if (ArrowLenMapRev.TryGetValue(value, out var ooxmlVal))
            endElem.SetAttributeValue("len", ooxmlVal);
        Save();
    }

    private static readonly string[] LnChildOrder =
    [
        "noFill", "solidFill", "gradFill", "pattFill",
        "prstDash", "custDash",
        "round", "bevel", "miter",
        "headEnd", "tailEnd", "extLst",
    ];

    /// <summary>
    /// Inserts a child element into the line element at the correct OOXML position.
    /// </summary>
    internal static XElement InsertLnChild(XElement ln, XName tag, params (string name, string value)[] attribs)
    {
        var newEl = new XElement(tag);
        foreach (var (name, value) in attribs)
            newEl.SetAttributeValue(name, value);

        var localName = tag.LocalName;
        var newRank = Array.IndexOf(LnChildOrder, localName);
        if (newRank < 0)
        {
            ln.Add(newEl);
            return newEl;
        }

        foreach (var child in ln.Elements().ToList())
        {
            var childRank = Array.IndexOf(LnChildOrder, child.Name.LocalName);
            if (childRank >= 0 && childRank > newRank)
            {
                child.AddBeforeSelf(newEl);
                return newEl;
            }
        }

        ln.Add(newEl);
        return newEl;
    }
}
