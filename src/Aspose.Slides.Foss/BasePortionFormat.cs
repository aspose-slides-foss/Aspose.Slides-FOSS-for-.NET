using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Common text-run formatting properties backed by an OOXML &lt;a:rPr&gt; element.
/// </summary>
public class BasePortionFormat : PVIObject, IBasePortionFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XName AHighlight = ANs + "highlight";
    private static readonly XName AULn = ANs + "uLn";
    private static readonly XName AUFill = ANs + "uFill";
    private static readonly XName ALatin = ANs + "latin";
    private static readonly XName AEa = ANs + "ea";
    private static readonly XName ACs = ANs + "cs";
    private static readonly XName ASym = ANs + "sym";

    #region Enum Maps

    private static readonly Dictionary<string, TextUnderlineType> UnderlineMap = new()
    {
        ["sng"] = TextUnderlineType.Single,
        ["dbl"] = TextUnderlineType.Double,
        ["heavy"] = TextUnderlineType.Heavy,
        ["dotted"] = TextUnderlineType.Dotted,
        ["dottedHeavy"] = TextUnderlineType.HeavyDotted,
        ["dash"] = TextUnderlineType.Dashed,
        ["dashHeavy"] = TextUnderlineType.HeavyDashed,
        ["dashLong"] = TextUnderlineType.LongDashed,
        ["dashLongHeavy"] = TextUnderlineType.HeavyLongDashed,
        ["dotDash"] = TextUnderlineType.DotDash,
        ["dotDashHeavy"] = TextUnderlineType.HeavyDotDash,
        ["dotDotDash"] = TextUnderlineType.DotDotDash,
        ["dotDotDashHeavy"] = TextUnderlineType.HeavyDotDotDash,
        ["wavy"] = TextUnderlineType.Wavy,
        ["wavyHeavy"] = TextUnderlineType.HeavyWavy,
        ["wavyDbl"] = TextUnderlineType.DoubleWavy,
        ["words"] = TextUnderlineType.Words,
        ["none"] = TextUnderlineType.None,
    };

    private static readonly Dictionary<TextUnderlineType, string> UnderlineMapRev =
        UnderlineMap.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private static readonly Dictionary<string, TextCapType> CapMap = new()
    {
        ["none"] = TextCapType.None,
        ["small"] = TextCapType.Small,
        ["all"] = TextCapType.All,
    };

    private static readonly Dictionary<TextCapType, string> CapMapRev =
        CapMap.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private static readonly Dictionary<string, TextStrikethroughType> StrikeMap = new()
    {
        ["noStrike"] = TextStrikethroughType.None,
        ["sngStrike"] = TextStrikethroughType.Single,
        ["dblStrike"] = TextStrikethroughType.Double,
    };

    private static readonly Dictionary<TextStrikethroughType, string> StrikeMapRev =
        StrikeMap.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    #endregion

    private XElement? _rprElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Gets the backing &lt;a:rPr&gt; element for internal use.
    /// </summary>
    internal XElement? RprElement => _rprElement;

    /// <summary>
    /// Initializes a new detached instance with a fresh &lt;a:rPr&gt; element.
    /// </summary>
    public BasePortionFormat()
    {
        _rprElement = new XElement(ANs + "rPr");
        _slidePart = null;
        _parentSlide = null;
    }

    /// <summary>
    /// Replaces the detached element with a live one from a parsed document.
    /// </summary>
    /// <param name="rprElement">The &lt;a:rPr&gt; XML element from the document.</param>
    /// <param name="slidePart">The slide part for serialization.</param>
    /// <param name="parentSlide">The parent slide reference.</param>
    internal void InitInternal(XElement? rprElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _rprElement = rprElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private void Save()
    {
        _slidePart?.Save();
    }

    #region NullableBool Attribute Helpers

    private NullableBool GetNullableBoolAttr(string attr)
    {
        if (_rprElement is null)
            return NullableBool.NotDefined;

        var value = _rprElement.Attribute(attr)?.Value;
        if (value is null)
            return NullableBool.NotDefined;

        return value == "1" ? NullableBool.True : NullableBool.False;
    }

    private void SetNullableBoolAttr(string attr, NullableBool value)
    {
        if (_rprElement is null)
            return;

        if (value == NullableBool.NotDefined)
        {
            _rprElement.Attribute(attr)?.Remove();
        }
        else
        {
            _rprElement.SetAttributeValue(attr, value == NullableBool.True ? "1" : "0");
        }

        Save();
    }

    #endregion

    #region Font Helpers

    private IFontData? GetFont(XName tag)
    {
        if (_rprElement is null)
            return null;

        var el = _rprElement.Element(tag);
        var typeface = el?.Attribute("typeface")?.Value;
        return typeface is not null ? new FontData(typeface) : null;
    }

    private void SetFont(XName tag, IFontData? value)
    {
        if (_rprElement is null)
            return;

        if (value is null)
        {
            _rprElement.Element(tag)?.Remove();
        }
        else
        {
            var el = _rprElement.Element(tag);
            if (el is null)
            {
                el = new XElement(tag);
                _rprElement.Add(el);
            }

            el.SetAttributeValue("typeface", value.FontName);
        }

        Save();
    }

    #endregion

    #region Read-Only Format-Object Properties

    /// <inheritdoc/>
    public ILineFormat? LineFormat
    {
        get
        {
            if (_rprElement is null)
                return null;

            var lf = new LineFormat();
            lf.InitInternal(_rprElement, _parentSlide);
            return lf;
        }
    }

    /// <inheritdoc/>
    public IFillFormat? FillFormat
    {
        get
        {
            if (_rprElement is null)
                return null;

            var ff = new FillFormat();
            ff.InitInternal(_rprElement, _parentSlide);
            return ff;
        }
    }

    /// <inheritdoc/>
    public IEffectFormat? EffectFormat
    {
        get
        {
            if (_rprElement is null)
                return null;

            var ef = new EffectFormat();
            ef.InitInternal(_rprElement, _parentSlide);
            return ef;
        }
    }

    /// <inheritdoc/>
    public IColorFormat? HighlightColor
    {
        get
        {
            if (_rprElement is null)
                return null;

            var highlight = _rprElement.Element(AHighlight);
            if (highlight is null)
            {
                highlight = new XElement(AHighlight);
                _rprElement.Add(highlight);
            }

            var cf = new ColorFormat();
            cf.InitInternal(highlight, _parentSlide);
            return cf;
        }
    }

    /// <inheritdoc/>
    public ILineFormat? UnderlineLineFormat
    {
        get
        {
            if (_rprElement is null)
                return null;

            var lf = new LineFormat();
            lf.InitInternal(_rprElement, _parentSlide, AULn);
            return lf;
        }
    }

    /// <inheritdoc/>
    public IFillFormat? UnderlineFillFormat
    {
        get
        {
            if (_rprElement is null)
                return null;

            var uFill = _rprElement.Element(AUFill);
            if (uFill is null)
            {
                uFill = new XElement(AUFill);
                _rprElement.Add(uFill);
            }

            var ff = new FillFormat();
            ff.InitInternal(uFill, _parentSlide);
            return ff;
        }
    }

    #endregion

    #region NullableBool Attribute Properties

    /// <inheritdoc/>
    public NullableBool FontBold
    {
        get => GetNullableBoolAttr("b");
        set => SetNullableBoolAttr("b", value);
    }

    /// <inheritdoc/>
    public NullableBool FontItalic
    {
        get => GetNullableBoolAttr("i");
        set => SetNullableBoolAttr("i", value);
    }

    /// <inheritdoc/>
    public NullableBool Kumimoji
    {
        get => GetNullableBoolAttr("kumimoji");
        set => SetNullableBoolAttr("kumimoji", value);
    }

    /// <inheritdoc/>
    public NullableBool NormaliseHeight
    {
        get => GetNullableBoolAttr("normalizeH");
        set => SetNullableBoolAttr("normalizeH", value);
    }

    /// <inheritdoc/>
    public NullableBool ProofDisabled
    {
        get => GetNullableBoolAttr("noProof");
        set => SetNullableBoolAttr("noProof", value);
    }

    #endregion

    #region Enum-Mapped Attribute Properties

    /// <inheritdoc/>
    public TextUnderlineType FontUnderline
    {
        get
        {
            if (_rprElement is null)
                return TextUnderlineType.NotDefined;

            var value = _rprElement.Attribute("u")?.Value;
            if (value is null)
                return TextUnderlineType.NotDefined;

            return UnderlineMap.TryGetValue(value, out var mapped) ? mapped : TextUnderlineType.NotDefined;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (value == TextUnderlineType.NotDefined)
            {
                _rprElement.Attribute("u")?.Remove();
            }
            else if (UnderlineMapRev.TryGetValue(value, out var xmlValue))
            {
                _rprElement.SetAttributeValue("u", xmlValue);
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public TextCapType TextCapType
    {
        get
        {
            if (_rprElement is null)
                return TextCapType.NotDefined;

            var value = _rprElement.Attribute("cap")?.Value;
            if (value is null)
                return TextCapType.NotDefined;

            return CapMap.TryGetValue(value, out var mapped) ? mapped : TextCapType.NotDefined;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (value == TextCapType.NotDefined)
            {
                _rprElement.Attribute("cap")?.Remove();
            }
            else if (CapMapRev.TryGetValue(value, out var xmlValue))
            {
                _rprElement.SetAttributeValue("cap", xmlValue);
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public TextStrikethroughType StrikethroughType
    {
        get
        {
            if (_rprElement is null)
                return TextStrikethroughType.NotDefined;

            var value = _rprElement.Attribute("strike")?.Value;
            if (value is null)
                return TextStrikethroughType.NotDefined;

            return StrikeMap.TryGetValue(value, out var mapped) ? mapped : TextStrikethroughType.NotDefined;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (value == TextStrikethroughType.NotDefined)
            {
                _rprElement.Attribute("strike")?.Remove();
            }
            else if (StrikeMapRev.TryGetValue(value, out var xmlValue))
            {
                _rprElement.SetAttributeValue("strike", xmlValue);
            }

            Save();
        }
    }

    #endregion

    #region NullableBool Element-Based Properties

    /// <inheritdoc/>
    public NullableBool IsHardUnderlineLine
    {
        get
        {
            if (_rprElement is null)
                return NullableBool.NotDefined;

            return _rprElement.Element(AULn) is not null ? NullableBool.True : NullableBool.False;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (value == NullableBool.True)
            {
                if (_rprElement.Element(AULn) is null)
                    _rprElement.Add(new XElement(AULn));
            }
            else
            {
                _rprElement.Element(AULn)?.Remove();
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public NullableBool IsHardUnderlineFill
    {
        get
        {
            if (_rprElement is null)
                return NullableBool.NotDefined;

            return _rprElement.Element(AUFill) is not null ? NullableBool.True : NullableBool.False;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (value == NullableBool.True)
            {
                if (_rprElement.Element(AUFill) is null)
                    _rprElement.Add(new XElement(AUFill));
            }
            else
            {
                _rprElement.Element(AUFill)?.Remove();
            }

            Save();
        }
    }

    #endregion

    #region Float Attribute Properties

    /// <inheritdoc/>
    public float FontHeight
    {
        get
        {
            if (_rprElement is null)
                return float.NaN;

            var value = _rprElement.Attribute("sz")?.Value;
            return value is null ? float.NaN : int.Parse(value) / 100.0f;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (float.IsNaN(value))
            {
                _rprElement.Attribute("sz")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("sz", ((int)Math.Round(value * 100)).ToString());
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public float Escapement
    {
        get
        {
            if (_rprElement is null)
                return float.NaN;

            var value = _rprElement.Attribute("baseline")?.Value;
            return value is null ? float.NaN : float.Parse(value) / 1000.0f;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (float.IsNaN(value))
            {
                _rprElement.Attribute("baseline")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("baseline", ((int)Math.Round(value * 1000)).ToString());
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public float KerningMinimalSize
    {
        get
        {
            if (_rprElement is null)
                return float.NaN;

            var value = _rprElement.Attribute("kern")?.Value;
            return value is null ? float.NaN : int.Parse(value) / 100.0f;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (float.IsNaN(value))
            {
                _rprElement.Attribute("kern")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("kern", ((int)Math.Round(value * 100)).ToString());
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public float Spacing
    {
        get
        {
            if (_rprElement is null)
                return float.NaN;

            var value = _rprElement.Attribute("spc")?.Value;
            return value is null ? float.NaN : int.Parse(value) / 100.0f;
        }
        set
        {
            if (_rprElement is null)
                return;

            if (float.IsNaN(value))
            {
                _rprElement.Attribute("spc")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("spc", ((int)Math.Round(value * 100)).ToString());
            }

            Save();
        }
    }

    #endregion

    #region Font Properties

    /// <inheritdoc/>
    public IFontData? LatinFont
    {
        get => GetFont(ALatin);
        set => SetFont(ALatin, value);
    }

    /// <inheritdoc/>
    public IFontData? EastAsianFont
    {
        get => GetFont(AEa);
        set => SetFont(AEa, value);
    }

    /// <inheritdoc/>
    public IFontData? ComplexScriptFont
    {
        get => GetFont(ACs);
        set => SetFont(ACs, value);
    }

    /// <inheritdoc/>
    public IFontData? SymbolFont
    {
        get => GetFont(ASym);
        set => SetFont(ASym, value);
    }

    #endregion

    #region String Attribute Properties

    /// <inheritdoc/>
    public string? LanguageId
    {
        get => _rprElement?.Attribute("lang")?.Value;
        set
        {
            if (_rprElement is null)
                return;

            if (string.IsNullOrEmpty(value))
            {
                _rprElement.Attribute("lang")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("lang", value);
            }

            Save();
        }
    }

    /// <inheritdoc/>
    public string? AlternativeLanguageId
    {
        get => _rprElement?.Attribute("altLang")?.Value;
        set
        {
            if (_rprElement is null)
                return;

            if (string.IsNullOrEmpty(value))
            {
                _rprElement.Attribute("altLang")?.Remove();
            }
            else
            {
                _rprElement.SetAttributeValue("altLang", value);
            }

            Save();
        }
    }

    #endregion

    #region SpellCheck

    /// <inheritdoc/>
    public NullableBool SpellCheck
    {
        get => GetNullableBoolAttr("err");
        set => SetNullableBoolAttr("err", value);
    }

    #endregion
}
