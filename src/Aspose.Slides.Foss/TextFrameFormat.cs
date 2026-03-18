using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Contains the TextFrame's formatting properties.
/// </summary>
public sealed class TextFrameFormat : ITextFrameFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const float EmuPerPoint = 12700f;
    private const int RotationUnit = 60000;

    // Default margins in EMU: left/right = 91440 (0.1 in), top/bottom = 45720 (0.05 in)
    private const int DefaultLRMarginEmu = 91440;
    private const int DefaultTBMarginEmu = 45720;

    private static readonly Dictionary<string, TextAnchorType> AnchorMap = new()
    {
        ["t"] = TextAnchorType.Top,
        ["ctr"] = TextAnchorType.Center,
        ["b"] = TextAnchorType.Bottom,
        ["just"] = TextAnchorType.Justified,
        ["dist"] = TextAnchorType.Distributed,
    };

    private static readonly Dictionary<TextAnchorType, string> AnchorMapRev =
        AnchorMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, TextVerticalType> VertMap = new()
    {
        ["horz"] = TextVerticalType.Horizontal,
        ["vert"] = TextVerticalType.Vertical,
        ["vert270"] = TextVerticalType.Vertical270,
        ["wordArtVert"] = TextVerticalType.WordArtVertical,
        ["eaVert"] = TextVerticalType.EastAsianVertical,
        ["mongolianVert"] = TextVerticalType.MongolianVertical,
        ["wordArtVertRtl"] = TextVerticalType.WordArtVerticalRightToLeft,
    };

    private static readonly Dictionary<TextVerticalType, string> VertMapRev =
        VertMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, TextShapeType> WarpMap = new()
    {
        ["textNoShape"] = TextShapeType.None,
        ["textPlain"] = TextShapeType.Plain,
        ["textStop"] = TextShapeType.Stop,
        ["textTriangle"] = TextShapeType.Triangle,
        ["textTriangleInverted"] = TextShapeType.TriangleInverted,
        ["textChevron"] = TextShapeType.Chevron,
        ["textChevronInverted"] = TextShapeType.ChevronInverted,
        ["textRingInside"] = TextShapeType.RingInside,
        ["textRingOutside"] = TextShapeType.RingOutside,
        ["textArchUp"] = TextShapeType.ArchUp,
        ["textArchDown"] = TextShapeType.ArchDown,
        ["textCircle"] = TextShapeType.Circle,
        ["textButton"] = TextShapeType.Button,
        ["textArchUpPour"] = TextShapeType.ArchUpPour,
        ["textArchDownPour"] = TextShapeType.ArchDownPour,
        ["textCirclePour"] = TextShapeType.CirclePour,
        ["textButtonPour"] = TextShapeType.ButtonPour,
        ["textCurveUp"] = TextShapeType.CurveUp,
        ["textCurveDown"] = TextShapeType.CurveDown,
        ["textCanUp"] = TextShapeType.CanUp,
        ["textCanDown"] = TextShapeType.CanDown,
        ["textWave1"] = TextShapeType.Wave1,
        ["textWave2"] = TextShapeType.Wave2,
        ["textDoubleWave1"] = TextShapeType.DoubleWave1,
        ["textWave4"] = TextShapeType.Wave4,
        ["textInflate"] = TextShapeType.Inflate,
        ["textDeflate"] = TextShapeType.Deflate,
        ["textInflateBottom"] = TextShapeType.InflateBottom,
        ["textDeflateBottom"] = TextShapeType.DeflateBottom,
        ["textInflateTop"] = TextShapeType.InflateTop,
        ["textDeflateTop"] = TextShapeType.DeflateTop,
        ["textDeflateInflate"] = TextShapeType.DeflateInflate,
        ["textDeflateInflateDeflate"] = TextShapeType.DeflateInflateDeflate,
        ["textFadeRight"] = TextShapeType.FadeRight,
        ["textFadeLeft"] = TextShapeType.FadeLeft,
        ["textFadeUp"] = TextShapeType.FadeUp,
        ["textFadeDown"] = TextShapeType.FadeDown,
        ["textSlantUp"] = TextShapeType.SlantUp,
        ["textSlantDown"] = TextShapeType.SlantDown,
        ["textCascadeUp"] = TextShapeType.CascadeUp,
        ["textCascadeDown"] = TextShapeType.CascadeDown,
    };

    private static readonly Dictionary<TextShapeType, string> WarpMapRev =
        WarpMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private static readonly Dictionary<string, TextAutofitType> AutofitMap = new()
    {
        ["noAutofit"] = TextAutofitType.None,
        ["normAutofit"] = TextAutofitType.Normal,
        ["spAutoFit"] = TextAutofitType.Shape,
    };

    private static readonly Dictionary<TextAutofitType, string> AutofitMapRev =
        AutofitMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    private XElement _txBodyElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Gets the backing &lt;a:txBody&gt; element for internal use.
    /// </summary>
    internal XElement TxBodyElement => _txBodyElement;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes a new instance of <see cref="TextFrameFormat"/> with a detached body.
    /// </summary>
    public TextFrameFormat()
    {
        _txBodyElement = new XElement(ANs + "txBody",
            new XElement(ANs + "bodyPr"));
    }

    /// <summary>
    /// Initializes internal state binding this format to real XML and a slide part.
    /// </summary>
    internal void InitInternal(XElement txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _txBodyElement = txBodyElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private XElement? GetBodyPr() => _txBodyElement?.Element(ANs + "bodyPr");

    private XElement EnsureBodyPr()
    {
        var bodyPr = GetBodyPr();
        if (bodyPr is not null) return bodyPr;
        bodyPr = new XElement(ANs + "bodyPr");
        _txBodyElement.Add(bodyPr);
        return bodyPr;
    }

    private void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public IPresentationComponent AsIPresentationComponent => new PresentationComponentWrapper(_parentSlide);

    // --- Margin properties (EMU <-> points) ---

    private float GetMargin(string attr, int defaultEmu)
    {
        var bodyPr = GetBodyPr();
        if (bodyPr is null) return defaultEmu / EmuPerPoint;
        var val = bodyPr.Attribute(attr)?.Value;
        if (val is null) return defaultEmu / EmuPerPoint;
        return int.TryParse(val, out var v) ? v / EmuPerPoint : defaultEmu / EmuPerPoint;
    }

    private void SetMargin(string attr, float value)
    {
        var bodyPr = EnsureBodyPr();
        bodyPr.SetAttributeValue(attr, (int)Math.Round(value * EmuPerPoint));
        Save();
    }

    /// <inheritdoc/>
    public float MarginLeft
    {
        get => GetMargin("lIns", DefaultLRMarginEmu);
        set => SetMargin("lIns", value);
    }

    /// <inheritdoc/>
    public float MarginRight
    {
        get => GetMargin("rIns", DefaultLRMarginEmu);
        set => SetMargin("rIns", value);
    }

    /// <inheritdoc/>
    public float MarginTop
    {
        get => GetMargin("tIns", DefaultTBMarginEmu);
        set => SetMargin("tIns", value);
    }

    /// <inheritdoc/>
    public float MarginBottom
    {
        get => GetMargin("bIns", DefaultTBMarginEmu);
        set => SetMargin("bIns", value);
    }

    /// <inheritdoc/>
    public NullableBool WrapText
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return NullableBool.NotDefined;
            var val = bodyPr.Attribute("wrap")?.Value;
            return val switch
            {
                "square" => NullableBool.True,
                "none" => NullableBool.False,
                _ => NullableBool.NotDefined,
            };
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            if (value == NullableBool.True)
                bodyPr.SetAttributeValue("wrap", "square");
            else if (value == NullableBool.False)
                bodyPr.SetAttributeValue("wrap", "none");
            else
                bodyPr.Attribute("wrap")?.Remove();
            Save();
        }
    }

    /// <inheritdoc/>
    public TextAnchorType AnchoringType
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return TextAnchorType.NotDefined;
            var val = bodyPr.Attribute("anchor")?.Value;
            if (val is not null && AnchorMap.TryGetValue(val, out var type))
                return type;
            return TextAnchorType.NotDefined;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            if (value == TextAnchorType.NotDefined)
                bodyPr.Attribute("anchor")?.Remove();
            else if (AnchorMapRev.TryGetValue(value, out var ooxmlVal))
                bodyPr.SetAttributeValue("anchor", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public NullableBool CenterText
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return NullableBool.NotDefined;
            var val = bodyPr.Attribute("anchorCtr")?.Value;
            return val switch
            {
                "1" or "true" => NullableBool.True,
                "0" or "false" => NullableBool.False,
                _ => NullableBool.NotDefined,
            };
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            if (value == NullableBool.True)
                bodyPr.SetAttributeValue("anchorCtr", "1");
            else if (value == NullableBool.False)
                bodyPr.SetAttributeValue("anchorCtr", "0");
            else
                bodyPr.Attribute("anchorCtr")?.Remove();
            Save();
        }
    }

    /// <inheritdoc/>
    public TextVerticalType TextVerticalType
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return TextVerticalType.NotDefined;
            var val = bodyPr.Attribute("vert")?.Value;
            if (val is not null && VertMap.TryGetValue(val, out var type))
                return type;
            return TextVerticalType.NotDefined;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            if (value == TextVerticalType.NotDefined)
                bodyPr.Attribute("vert")?.Remove();
            else if (VertMapRev.TryGetValue(value, out var ooxmlVal))
                bodyPr.SetAttributeValue("vert", ooxmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public TextAutofitType AutofitType
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return TextAutofitType.NotDefined;
            foreach (var kvp in AutofitMap)
            {
                if (bodyPr.Element(ANs + kvp.Key) is not null)
                    return kvp.Value;
            }
            return TextAutofitType.NotDefined;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            // Remove existing autofit elements
            foreach (var tag in AutofitMap.Keys)
                bodyPr.Element(ANs + tag)?.Remove();

            if (value == TextAutofitType.Shape)
            {
                bodyPr.Add(new XElement(ANs + "spAutoFit"));
                ResizeShapeToFitText(bodyPr);
            }
            else if (value != TextAutofitType.NotDefined && AutofitMapRev.TryGetValue(value, out var tag2))
            {
                bodyPr.Add(new XElement(ANs + tag2));
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public int ColumnCount
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return 1;
            var val = bodyPr.Attribute("numCol")?.Value;
            if (val is not null && int.TryParse(val, out var n) && n > 0)
                return n;
            return 1;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            bodyPr.SetAttributeValue("numCol", value);
            Save();
        }
    }

    /// <inheritdoc/>
    public float ColumnSpacing
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return 0f;
            var val = bodyPr.Attribute("spcCol")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            bodyPr.SetAttributeValue("spcCol", (int)Math.Round(value * EmuPerPoint));
            Save();
        }
    }

    /// <inheritdoc/>
    public float RotationAngle
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return 0f;
            var val = bodyPr.Attribute("rot")?.Value;
            return val is not null && int.TryParse(val, out var v) ? (float)v / RotationUnit : 0f;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            bodyPr.SetAttributeValue("rot", (int)Math.Round(value * RotationUnit));
            Save();
        }
    }

    /// <inheritdoc/>
    public TextShapeType Transform
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return TextShapeType.NotDefined;
            var prstTxWarp = bodyPr.Element(ANs + "prstTxWarp");
            if (prstTxWarp is null) return TextShapeType.NotDefined;
            var prst = prstTxWarp.Attribute("prst")?.Value;
            if (prst is not null && WarpMap.TryGetValue(prst, out var type))
                return type;
            return TextShapeType.NotDefined;
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            bodyPr.Element(ANs + "prstTxWarp")?.Remove();
            if (value != TextShapeType.NotDefined && WarpMapRev.TryGetValue(value, out var ooxmlVal))
            {
                bodyPr.Add(new XElement(ANs + "prstTxWarp",
                    new XAttribute("prst", ooxmlVal)));
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public bool KeepTextFlat
    {
        get
        {
            var bodyPr = GetBodyPr();
            if (bodyPr is null) return false;
            var val = bodyPr.Attribute("flatTx")?.Value;
            return val is "1" or "true";
        }
        set
        {
            var bodyPr = EnsureBodyPr();
            bodyPr.SetAttributeValue("flatTx", value ? "1" : "0");
            Save();
        }
    }

    /// <inheritdoc/>
    public IThreeDFormat ThreeDFormat
    {
        get
        {
            var bodyPr = EnsureBodyPr();
            var tdf = new ThreeDFormat();
            tdf.InitInternal(bodyPr, _slidePart, _parentSlide);
            return tdf;
        }
    }

    /// <summary>
    /// Resizes the parent shape to fit text content, preserving vertical center.
    /// </summary>
    internal void ResizeShapeToFitText(XElement bodyPr)
    {
        var txBody = _txBodyElement;
        if (txBody is null) return;

        var spElement = txBody.Parent;
        if (spElement is null) return;

        var spPr = spElement.Element(PNs + "spPr");
        if (spPr is null) return;

        var xfrm = spPr.Element(ANs + "xfrm");
        if (xfrm is null) return;

        var ext = xfrm.Element(ANs + "ext");
        var off = xfrm.Element(ANs + "off");
        if (ext is null || off is null) return;

        // Count paragraphs
        int numParagraphs = Math.Max(1, txBody.Elements(ANs + "p").Count());

        // Default font size 18pt, single line spacing = 120% of font size
        int defaultFontSizeEmu = (int)(18.0 * EmuPerPoint);
        int lineHeightEmu = (int)(defaultFontSizeEmu * 1.2);
        int textHeightEmu = numParagraphs * lineHeightEmu;

        // Top and bottom margins (default 45720 EMU = 3.6pt each)
        int topMargin = int.TryParse(bodyPr.Attribute("tIns")?.Value, out var t) ? t : 45720;
        int bottomMargin = int.TryParse(bodyPr.Attribute("bIns")?.Value, out var b) ? b : 45720;
        int requiredHeight = textHeightEmu + topMargin + bottomMargin;

        // Resize preserving vertical center
        int currentY = int.TryParse(off.Attribute("y")?.Value, out var y) ? y : 0;
        int currentCy = int.TryParse(ext.Attribute("cy")?.Value, out var cy) ? cy : 0;
        int centerY = currentY + currentCy / 2;
        int newY = centerY - requiredHeight / 2;

        ext.SetAttributeValue("cy", requiredHeight);
        off.SetAttributeValue("y", newY);
    }

    /// <summary>
    /// Lightweight wrapper that delegates <see cref="IPresentationComponent.Presentation"/> to a slide.
    /// </summary>
    private sealed class PresentationComponentWrapper(IBaseSlide? parentSlide) : IPresentationComponent
    {
        /// <inheritdoc/>
        public override IPresentation? Presentation => parentSlide?.Presentation;
    }
}
