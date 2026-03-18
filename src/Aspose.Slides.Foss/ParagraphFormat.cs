using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the formatting properties for a paragraph.
/// Unlike <see cref="IParagraph"/>, all properties of this class are writeable.
/// </summary>
public sealed class ParagraphFormat : PVIObject, IParagraphFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const float EmuPerPoint = 12700f;

    #region Enum Maps

    private static readonly Dictionary<string, TextAlignment> AlignmentMap = new()
    {
        ["l"] = TextAlignment.Left,
        ["ctr"] = TextAlignment.Center,
        ["r"] = TextAlignment.Right,
        ["just"] = TextAlignment.Justify,
        ["justLow"] = TextAlignment.JustifyLow,
        ["dist"] = TextAlignment.Distributed,
    };

    private static readonly Dictionary<TextAlignment, string> AlignmentMapRev =
        AlignmentMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, FontAlignment> FontAlignMap = new()
    {
        ["auto"] = FontAlignment.Automatic,
        ["t"] = FontAlignment.Top,
        ["ctr"] = FontAlignment.Center,
        ["b"] = FontAlignment.Bottom,
        ["base"] = FontAlignment.Baseline,
    };

    private static readonly Dictionary<FontAlignment, string> FontAlignMapRev =
        FontAlignMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    #endregion

    #region Child Element Ordering

    // OOXML schema order for <a:pPr> child elements (CT_TextParagraphProperties).
    private static readonly string[][] PprChildOrder =
    [
        ["lnSpc"],
        ["spcBef"],
        ["spcAft"],
        ["buClrTx", "buClr"],
        ["buSzTx", "buSzPct", "buSzPts"],
        ["buFontTx", "buFont"],
        ["buNone", "buAutoNum", "buChar", "buBlip"],
        ["tabLst"],
        ["defRPr"],
        ["extLst"],
    ];

    private static readonly Dictionary<string, int> PprTagIndex = BuildTagIndex();

    private static Dictionary<string, int> BuildTagIndex()
    {
        var index = new Dictionary<string, int>();
        for (var i = 0; i < PprChildOrder.Length; i++)
        {
            foreach (var tag in PprChildOrder[i])
                index[tag] = i;
        }
        return index;
    }

    #endregion

    private XElement? _pprElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Gets the backing &lt;a:pPr&gt; element for internal use.
    /// </summary>
    internal XElement? PprElement => _pprElement;
    private readonly BulletFormat _bullet = new();
    private readonly PortionFormat _defaultPortionFormat = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="ParagraphFormat"/> class
    /// with a detached &lt;a:pPr&gt; element.
    /// </summary>
    public ParagraphFormat()
    {
        _pprElement = new XElement(ANs + "pPr");
    }

    /// <summary>
    /// Initializes internal state from an existing XML element.
    /// </summary>
    internal void InitInternal(XElement pprElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _pprElement = pprElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _bullet.InitInternal(pprElement, slidePart, parentSlide);

        var defRPr = pprElement.Element(ANs + "defRPr");
        if (defRPr is not null)
            _defaultPortionFormat.InitInternal(defRPr, slidePart, parentSlide);
    }

    private void Save() => _slidePart?.Save();

    /// <summary>
    /// Inserts a child element into the &lt;a:pPr&gt; at the correct schema position.
    /// </summary>
    private XElement InsertPprChild(string localName, params (string name, string value)[] attribs)
    {
        var tag = ANs + localName;
        var el = new XElement(tag);
        foreach (var (name, value) in attribs)
            el.SetAttributeValue(name, value);

        if (_pprElement is null) return el;

        var targetPos = PprTagIndex.GetValueOrDefault(localName, 999);
        XElement? insertBefore = null;
        foreach (var child in _pprElement.Elements())
        {
            var childPos = PprTagIndex.GetValueOrDefault(child.Name.LocalName, 999);
            if (childPos > targetPos)
            {
                insertBefore = child;
                break;
            }
        }

        if (insertBefore is not null)
            insertBefore.AddBeforeSelf(el);
        else
            _pprElement.Add(el);

        return el;
    }

    #region NullableBool Helpers

    private NullableBool GetNullableBoolAttr(string attr)
    {
        if (_pprElement is null) return NullableBool.NotDefined;
        var val = _pprElement.Attribute(attr)?.Value;
        return val switch
        {
            "1" => NullableBool.True,
            "0" => NullableBool.False,
            _ => NullableBool.NotDefined,
        };
    }

    private void SetNullableBoolAttr(string attr, NullableBool value)
    {
        if (_pprElement is null) return;
        if (value == NullableBool.NotDefined)
            _pprElement.Attribute(attr)?.Remove();
        else
            _pprElement.SetAttributeValue(attr, value == NullableBool.True ? "1" : "0");
        Save();
    }

    #endregion

    #region Spacing Helpers

    private float GetSpacing(string localName)
    {
        if (_pprElement is null) return float.NaN;
        var el = _pprElement.Element(ANs + localName);
        if (el is null) return float.NaN;

        var pct = el.Element(ANs + "spcPct");
        if (pct is not null)
        {
            var val = pct.Attribute("val")?.Value;
            if (val is not null && int.TryParse(val, out var v))
                return v / 1000f;
        }

        var pts = el.Element(ANs + "spcPts");
        if (pts is not null)
        {
            var val = pts.Attribute("val")?.Value;
            if (val is not null && int.TryParse(val, out var v))
                return -(v / 100f);
        }

        return float.NaN;
    }

    private void SetSpacing(string localName, float value)
    {
        if (_pprElement is null) return;
        var el = _pprElement.Element(ANs + localName);

        if (float.IsNaN(value))
        {
            el?.Remove();
        }
        else
        {
            if (el is null)
                el = InsertPprChild(localName);
            else
                el.RemoveAll();

            if (value >= 0)
                el.Add(new XElement(ANs + "spcPct", new XAttribute("val", (int)Math.Round(value * 1000))));
            else
                el.Add(new XElement(ANs + "spcPts", new XAttribute("val", (int)Math.Round(-value * 100))));
        }

        Save();
    }

    #endregion

    #region EMU Attribute Helpers

    private float GetEmuAttr(string attr)
    {
        if (_pprElement is null) return float.NaN;
        var val = _pprElement.Attribute(attr)?.Value;
        if (val is null) return float.NaN;
        return int.TryParse(val, out var v) ? v / EmuPerPoint : float.NaN;
    }

    private void SetEmuAttr(string attr, float value)
    {
        if (_pprElement is null) return;
        if (float.IsNaN(value))
            _pprElement.Attribute(attr)?.Remove();
        else
            _pprElement.SetAttributeValue(attr, (int)Math.Round(value * EmuPerPoint));
        Save();
    }

    #endregion

    /// <inheritdoc />
    public IBulletFormat Bullet => _bullet;

    /// <inheritdoc />
    public int Depth
    {
        get
        {
            if (_pprElement is null) return 0;
            var val = _pprElement.Attribute("lvl")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v : 0;
        }
        set
        {
            if (_pprElement is null) return;
            if (value == 0)
                _pprElement.Attribute("lvl")?.Remove();
            else
                _pprElement.SetAttributeValue("lvl", value);
            Save();
        }
    }

    /// <inheritdoc />
    public TextAlignment Alignment
    {
        get
        {
            if (_pprElement is null) return TextAlignment.NotDefined;
            var val = _pprElement.Attribute("algn")?.Value;
            if (val is not null && AlignmentMap.TryGetValue(val, out var alignment))
                return alignment;
            return TextAlignment.NotDefined;
        }
        set
        {
            if (_pprElement is null) return;
            if (value == TextAlignment.NotDefined)
                _pprElement.Attribute("algn")?.Remove();
            else if (AlignmentMapRev.TryGetValue(value, out var ooxmlVal))
                _pprElement.SetAttributeValue("algn", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc />
    public float SpaceWithin
    {
        get => GetSpacing("lnSpc");
        set => SetSpacing("lnSpc", value);
    }

    /// <inheritdoc />
    public float SpaceBefore
    {
        get => GetSpacing("spcBef");
        set => SetSpacing("spcBef", value);
    }

    /// <inheritdoc />
    public float SpaceAfter
    {
        get => GetSpacing("spcAft");
        set => SetSpacing("spcAft", value);
    }

    /// <inheritdoc />
    public NullableBool EastAsianLineBreak
    {
        get => GetNullableBoolAttr("eaLnBrk");
        set => SetNullableBoolAttr("eaLnBrk", value);
    }

    /// <inheritdoc />
    public NullableBool RightToLeft
    {
        get => GetNullableBoolAttr("rtl");
        set => SetNullableBoolAttr("rtl", value);
    }

    /// <inheritdoc />
    public NullableBool LatinLineBreak
    {
        get => GetNullableBoolAttr("latinLnBrk");
        set => SetNullableBoolAttr("latinLnBrk", value);
    }

    /// <inheritdoc />
    public NullableBool HangingPunctuation
    {
        get => GetNullableBoolAttr("hangingPunct");
        set => SetNullableBoolAttr("hangingPunct", value);
    }

    /// <inheritdoc />
    public float MarginLeft
    {
        get => GetEmuAttr("marL");
        set => SetEmuAttr("marL", value);
    }

    /// <inheritdoc />
    public float MarginRight
    {
        get => GetEmuAttr("marR");
        set => SetEmuAttr("marR", value);
    }

    /// <inheritdoc />
    public float Indent
    {
        get => GetEmuAttr("indent");
        set => SetEmuAttr("indent", value);
    }

    /// <inheritdoc />
    public float DefaultTabSize
    {
        get => GetEmuAttr("defTabSz");
        set => SetEmuAttr("defTabSz", value);
    }

    /// <inheritdoc />
    public FontAlignment FontAlignment
    {
        get
        {
            if (_pprElement is null) return FontAlignment.Default;
            var val = _pprElement.Attribute("fontAlgn")?.Value;
            if (val is not null && FontAlignMap.TryGetValue(val, out var fa))
                return fa;
            return FontAlignment.Default;
        }
        set
        {
            if (_pprElement is null) return;
            if (value == FontAlignment.Default)
                _pprElement.Attribute("fontAlgn")?.Remove();
            else if (FontAlignMapRev.TryGetValue(value, out var ooxmlVal))
                _pprElement.SetAttributeValue("fontAlgn", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc />
    public PortionFormat DefaultPortionFormat => _defaultPortionFormat;
}
