using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Manages paragraph bullet formatting backed by OOXML bullet elements.
/// </summary>
public sealed class BulletFormat : PVIObject, IBulletFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly Dictionary<string, NumberedBulletStyle> AutoNumMap = new()
    {
        ["alphaLcParenBoth"] = NumberedBulletStyle.BulletAlphaLCParenBoth,
        ["alphaUcParenBoth"] = NumberedBulletStyle.BulletAlphaUCParenBoth,
        ["alphaLcParenR"] = NumberedBulletStyle.BulletAlphaLCParenRight,
        ["alphaUcParenR"] = NumberedBulletStyle.BulletAlphaUCParenRight,
        ["alphaLcPeriod"] = NumberedBulletStyle.BulletAlphaLCPeriod,
        ["alphaUcPeriod"] = NumberedBulletStyle.BulletAlphaUCPeriod,
        ["arabicParenBoth"] = NumberedBulletStyle.BulletArabicParenBoth,
        ["arabicParenR"] = NumberedBulletStyle.BulletArabicParenRight,
        ["arabicPeriod"] = NumberedBulletStyle.BulletArabicPeriod,
        ["arabicPlain"] = NumberedBulletStyle.BulletArabicPlain,
        ["romanLcParenBoth"] = NumberedBulletStyle.BulletRomanLCParenBoth,
        ["romanUcParenBoth"] = NumberedBulletStyle.BulletRomanUCParenBoth,
        ["romanLcParenR"] = NumberedBulletStyle.BulletRomanLCParenRight,
        ["romanUcParenR"] = NumberedBulletStyle.BulletRomanUCParenRight,
        ["romanLcPeriod"] = NumberedBulletStyle.BulletRomanLCPeriod,
        ["romanUcPeriod"] = NumberedBulletStyle.BulletRomanUCPeriod,
        ["circleNumDbPlain"] = NumberedBulletStyle.BulletCircleNumDBPlain,
        ["circleNumWdBlackPlain"] = NumberedBulletStyle.BulletCircleNumWDBlackPlain,
        ["circleNumWdWhitePlain"] = NumberedBulletStyle.BulletCircleNumWDWhitePlain,
        ["ea1ChsPeriod"] = NumberedBulletStyle.BulletSimpChinPeriod,
        ["ea1ChsPlain"] = NumberedBulletStyle.BulletSimpChinPlain,
        ["ea1ChtPeriod"] = NumberedBulletStyle.BulletTradChinPeriod,
        ["ea1ChtPlain"] = NumberedBulletStyle.BulletTradChinPlain,
        ["ea1JpnChsDbPeriod"] = NumberedBulletStyle.BulletKanjiSimpChinDBPeriod,
        ["ea1JpnKorPeriod"] = NumberedBulletStyle.BulletKoreanPeriod,
        ["ea1JpnKorPlain"] = NumberedBulletStyle.BulletKoreanPlain,
        ["arabic1Minus"] = NumberedBulletStyle.BulletArabicAlphaDash,
        ["arabic2Minus"] = NumberedBulletStyle.BulletArabicAbjadDash,
        ["hebrew2Minus"] = NumberedBulletStyle.BulletHebrewDash,
        ["thaiAlphaPeriod"] = NumberedBulletStyle.BulletThaiAlphaPeriod,
        ["thaiAlphaParenR"] = NumberedBulletStyle.BulletThaiAlphaParenRight,
        ["thaiAlphaParenBoth"] = NumberedBulletStyle.BulletThaiAlphaParenBoth,
        ["thaiNumPeriod"] = NumberedBulletStyle.BulletThaiNumPeriod,
        ["thaiNumParenR"] = NumberedBulletStyle.BulletThaiNumParenRight,
        ["thaiNumParenBoth"] = NumberedBulletStyle.BulletThaiNumParenBoth,
        ["hindiAlphaPeriod"] = NumberedBulletStyle.BulletHindiAlphaPeriod,
        ["hindiNumPeriod"] = NumberedBulletStyle.BulletHindiNumPeriod,
        ["hindiNumParenR"] = NumberedBulletStyle.BulletHindiNumParenRight,
        ["hindiAlpha1Period"] = NumberedBulletStyle.BulletHindiAlpha1Period,
        ["arabicDbPeriod"] = NumberedBulletStyle.BulletArabicDBPeriod,
        ["arabicDbPlain"] = NumberedBulletStyle.BulletArabicDBPlain,
    };

    private static readonly Dictionary<NumberedBulletStyle, string> ReverseAutoNumMap =
        AutoNumMap.ToDictionary(kv => kv.Value, kv => kv.Key);

    #region Child Element Ordering

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

    private XElement? _pPrElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state.
    /// </summary>
    internal void InitInternal(XElement pPrElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _pPrElement = pPrElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private void Save() => _slidePart?.Save();

    /// <summary>
    /// Inserts a child element into the pPr at the correct schema position.
    /// </summary>
    private XElement InsertPprChild(string localName, params (string name, string value)[] attribs)
    {
        var el = new XElement(ANs + localName);
        foreach (var (name, value) in attribs)
            el.SetAttributeValue(name, value);

        if (_pPrElement is null) return el;

        var targetPos = PprTagIndex.GetValueOrDefault(localName, 999);
        XElement? insertBefore = null;
        foreach (var child in _pPrElement.Elements())
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
            _pPrElement.Add(el);

        return el;
    }

    /// <inheritdoc/>
    public BulletType Type
    {
        get
        {
            if (_pPrElement is null) return BulletType.NotDefined;
            if (_pPrElement.Element(ANs + "buNone") is not null) return BulletType.None;
            if (_pPrElement.Element(ANs + "buChar") is not null) return BulletType.Symbol;
            if (_pPrElement.Element(ANs + "buAutoNum") is not null) return BulletType.Numbered;
            if (_pPrElement.Element(ANs + "buBlip") is not null) return BulletType.Picture;
            return BulletType.NotDefined;
        }
        set
        {
            if (_pPrElement is null) return;
            RemoveBulletElements();
            switch (value)
            {
                case BulletType.None:
                    InsertPprChild("buNone");
                    break;
                case BulletType.Symbol:
                    InsertPprChild("buChar", ("char", "\u2022"));
                    break;
                case BulletType.Numbered:
                    InsertPprChild("buAutoNum", ("type", "arabicPeriod"));
                    break;
                case BulletType.Picture:
                    InsertPprChild("buBlip");
                    break;
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public string Char
    {
        get => _pPrElement?.Element(ANs + "buChar")?.Attribute("char")?.Value ?? string.Empty;
        set
        {
            if (_pPrElement is null) return;
            var buChar = _pPrElement.Element(ANs + "buChar");
            if (buChar is null)
            {
                RemoveBulletElements();
                buChar = InsertPprChild("buChar");
            }
            buChar.SetAttributeValue("char", value);
            Save();
        }
    }

    /// <inheritdoc/>
    public IFontData? Font
    {
        get
        {
            var typeface = _pPrElement?.Element(ANs + "buFont")?.Attribute("typeface")?.Value;
            return typeface is not null ? new FontData(typeface) : null;
        }
        set
        {
            if (_pPrElement is null) return;
            var buFont = _pPrElement.Element(ANs + "buFont");
            if (value is null)
            {
                buFont?.Remove();
            }
            else
            {
                if (buFont is null)
                    buFont = InsertPprChild("buFont");
                buFont.SetAttributeValue("typeface", value.FontName);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get
        {
            if (_pPrElement is null) return float.NaN;
            var buSzPct = _pPrElement.Element(ANs + "buSzPct");
            if (buSzPct is not null)
            {
                var val = buSzPct.Attribute("val")?.Value;
                if (val is not null && int.TryParse(val, out var v))
                    return v / 1000f;
            }
            var buSzPts = _pPrElement.Element(ANs + "buSzPts");
            if (buSzPts is not null)
            {
                var val = buSzPts.Attribute("val")?.Value;
                if (val is not null && int.TryParse(val, out var v))
                    return v / 100f;
            }
            return float.NaN;
        }
        set
        {
            if (_pPrElement is null) return;
            if (float.IsNaN(value))
            {
                _pPrElement.Element(ANs + "buSzPct")?.Remove();
                _pPrElement.Element(ANs + "buSzPts")?.Remove();
            }
            else
            {
                _pPrElement.Element(ANs + "buSzPts")?.Remove();
                var buSzPct = _pPrElement.Element(ANs + "buSzPct");
                if (buSzPct is null)
                    buSzPct = InsertPprChild("buSzPct");
                buSzPct.SetAttributeValue("val", (int)(value * 1000));
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IColorFormat Color
    {
        get
        {
            var buClr = _pPrElement?.Element(ANs + "buClr");
            var cf = new ColorFormat();
            if (buClr is not null)
                cf.InitInternal(buClr, _parentSlide);
            return cf;
        }
    }

    /// <inheritdoc/>
    public int NumberedBulletStartWith
    {
        get
        {
            var val = _pPrElement?.Element(ANs + "buAutoNum")?.Attribute("startAt")?.Value;
            return val is not null && int.TryParse(val, out var v) ? v : 1;
        }
        set
        {
            var buAutoNum = _pPrElement?.Element(ANs + "buAutoNum");
            if (buAutoNum is not null)
            {
                buAutoNum.SetAttributeValue("startAt", value);
                Save();
            }
        }
    }

    /// <inheritdoc/>
    public NumberedBulletStyle NumberedBulletStyle
    {
        get
        {
            var type = _pPrElement?.Element(ANs + "buAutoNum")?.Attribute("type")?.Value;
            if (type is not null && AutoNumMap.TryGetValue(type, out var style))
                return style;
            return NumberedBulletStyle.NotDefined;
        }
        set
        {
            var buAutoNum = _pPrElement?.Element(ANs + "buAutoNum");
            if (buAutoNum is not null && ReverseAutoNumMap.TryGetValue(value, out var xml))
            {
                buAutoNum.SetAttributeValue("type", xml);
                Save();
            }
        }
    }

    /// <inheritdoc/>
    public NullableBool IsBulletHardColor
    {
        get
        {
            if (_pPrElement is null) return NullableBool.NotDefined;
            if (_pPrElement.Element(ANs + "buClr") is not null) return NullableBool.True;
            if (_pPrElement.Element(ANs + "buClrTx") is not null) return NullableBool.False;
            return NullableBool.NotDefined;
        }
        set
        {
            if (_pPrElement is null) return;
            _pPrElement.Element(ANs + "buClr")?.Remove();
            _pPrElement.Element(ANs + "buClrTx")?.Remove();
            if (value == NullableBool.True)
                InsertPprChild("buClr");
            else if (value == NullableBool.False)
                InsertPprChild("buClrTx");
            Save();
        }
    }

    /// <inheritdoc/>
    public NullableBool IsBulletHardFont
    {
        get
        {
            if (_pPrElement is null) return NullableBool.NotDefined;
            if (_pPrElement.Element(ANs + "buFont") is not null) return NullableBool.True;
            if (_pPrElement.Element(ANs + "buFontTx") is not null) return NullableBool.False;
            return NullableBool.NotDefined;
        }
        set
        {
            if (_pPrElement is null) return;
            _pPrElement.Element(ANs + "buFont")?.Remove();
            _pPrElement.Element(ANs + "buFontTx")?.Remove();
            if (value == NullableBool.True)
                InsertPprChild("buFont");
            else if (value == NullableBool.False)
                InsertPprChild("buFontTx");
            Save();
        }
    }

    /// <inheritdoc/>
    public ISlidesPicture? Picture => null;

    private void RemoveBulletElements()
    {
        _pPrElement?.Element(ANs + "buNone")?.Remove();
        _pPrElement?.Element(ANs + "buChar")?.Remove();
        _pPrElement?.Element(ANs + "buAutoNum")?.Remove();
        _pPrElement?.Element(ANs + "buBlip")?.Remove();
    }
}
