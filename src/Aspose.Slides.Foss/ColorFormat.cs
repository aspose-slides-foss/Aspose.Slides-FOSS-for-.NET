using System.Xml.Linq;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a color format used in presentation elements.
/// </summary>
public sealed class ColorFormat : IColorFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly HashSet<XName> ColorElementTags =
    [
        ANs + "srgbClr",
        ANs + "schemeClr",
        ANs + "prstClr",
        ANs + "sysClr",
        ANs + "hlsClr",
        ANs + "scrgbClr",
    ];

    private static readonly Dictionary<string, SchemeColor> SchemeClrToEnum = new()
    {
        ["bg1"] = SchemeColor.Background1,
        ["bg2"] = SchemeColor.Background2,
        ["tx1"] = SchemeColor.Text1,
        ["tx2"] = SchemeColor.Text2,
        ["accent1"] = SchemeColor.Accent1,
        ["accent2"] = SchemeColor.Accent2,
        ["accent3"] = SchemeColor.Accent3,
        ["accent4"] = SchemeColor.Accent4,
        ["accent5"] = SchemeColor.Accent5,
        ["accent6"] = SchemeColor.Accent6,
        ["hlink"] = SchemeColor.Hyperlink,
        ["folHlink"] = SchemeColor.FollowedHyperlink,
        ["phClr"] = SchemeColor.StyleColor,
        ["dk1"] = SchemeColor.Dark1,
        ["dk2"] = SchemeColor.Dark2,
        ["lt1"] = SchemeColor.Light1,
        ["lt2"] = SchemeColor.Light2,
    };

    private static readonly Dictionary<SchemeColor, string> EnumToSchemeClr =
        SchemeClrToEnum.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _parentElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes internal state for this color format.
    /// </summary>
    internal void InitInternal(XElement parentElement, IBaseSlide? parentSlide, SlidePart? slidePart = null)
    {
        _parentElement = parentElement;
        _parentSlide = parentSlide;
        _slidePart = slidePart;
    }

    private XElement? FindColorElement()
    {
        if (_parentElement is null) return null;
        foreach (var child in _parentElement.Elements())
        {
            if (ColorElementTags.Contains(child.Name))
                return child;
        }
        return null;
    }

    private void ClearColorElements()
    {
        if (_parentElement is null) return;
        var toRemove = _parentElement.Elements().Where(e => ColorElementTags.Contains(e.Name)).ToList();
        foreach (var el in toRemove)
            el.Remove();
    }

    private void Save()
    {
        _slidePart?.Save();
    }

    private static int ReadAlpha(XElement colorElement)
    {
        var alpha = colorElement.Element(ANs + "alpha");
        if (alpha is null) return 255;
        var valAttr = alpha.Attribute("val")?.Value;
        if (valAttr is null || !int.TryParse(valAttr, out var val)) return 255;
        return (int)Math.Round(val * 255.0 / 100000.0);
    }

    private static void WriteAlpha(XElement colorElement, int alpha)
    {
        colorElement.Element(ANs + "alpha")?.Remove();
        if (alpha < 255)
        {
            var val = (int)Math.Round(alpha * 100000.0 / 255.0);
            colorElement.Add(new XElement(ANs + "alpha", new XAttribute("val", val)));
        }
    }

    private static string PascalToCamel(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        return char.ToLowerInvariant(name[0]) + name[1..];
    }

    private static string CamelToPascal(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        return char.ToUpperInvariant(name[0]) + name[1..];
    }

    /// <summary>
    /// Converts a camelCase string to UPPER_SNAKE_CASE.
    /// For example, "aliceBlue" becomes "ALICE_BLUE".
    /// </summary>
    internal static string CamelToUpperSnake(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        var result = new System.Text.StringBuilder(name.Length + 4);
        for (int i = 0; i < name.Length; i++)
        {
            var ch = name[i];
            if (char.IsUpper(ch) && i > 0)
                result.Append('_');
            result.Append(char.ToUpperInvariant(ch));
        }
        return result.ToString();
    }

    /// <summary>
    /// Converts an UPPER_SNAKE_CASE string to camelCase.
    /// For example, "ALICE_BLUE" becomes "aliceBlue".
    /// </summary>
    internal static string UpperSnakeToCamel(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        var parts = name.ToLowerInvariant().Split('_');
        if (parts.Length == 0) return name;
        var result = new System.Text.StringBuilder(name.Length);
        result.Append(parts[0]);
        for (int i = 1; i < parts.Length; i++)
        {
            if (parts[i].Length > 0)
            {
                result.Append(char.ToUpperInvariant(parts[i][0]));
                result.Append(parts[i][1..]);
            }
        }
        return result.ToString();
    }

    /// <inheritdoc/>
    public ColorType ColorType
    {
        get
        {
            var el = FindColorElement();
            if (el is null) return ColorType.NotDefined;
            var localName = el.Name.LocalName;
            return localName switch
            {
                "srgbClr" => ColorType.RGB,
                "schemeClr" => ColorType.Scheme,
                "prstClr" => ColorType.Preset,
                "sysClr" => ColorType.System,
                "hlsClr" => ColorType.HSL,
                "scrgbClr" => ColorType.RGBPercentage,
                _ => ColorType.NotDefined,
            };
        }
        set
        {
            if (value == ColorType) return;
            ClearColorElements();
            if (_parentElement is null) return;
            switch (value)
            {
                case ColorType.RGB:
                    _parentElement.Add(new XElement(ANs + "srgbClr", new XAttribute("val", "000000")));
                    break;
                case ColorType.Scheme:
                    _parentElement.Add(new XElement(ANs + "schemeClr", new XAttribute("val", "tx1")));
                    break;
                case ColorType.Preset:
                    _parentElement.Add(new XElement(ANs + "prstClr", new XAttribute("val", "black")));
                    break;
                case ColorType.System:
                    _parentElement.Add(new XElement(ANs + "sysClr", new XAttribute("val", "windowText")));
                    break;
                case ColorType.HSL:
                    _parentElement.Add(new XElement(ANs + "hlsClr",
                        new XAttribute("hue", "0"),
                        new XAttribute("sat", "0"),
                        new XAttribute("lum", "0")));
                    break;
                case ColorType.RGBPercentage:
                    _parentElement.Add(new XElement(ANs + "scrgbClr",
                        new XAttribute("r", "0"),
                        new XAttribute("g", "0"),
                        new XAttribute("b", "0")));
                    break;
                case ColorType.NotDefined:
                    // Already cleared above
                    break;
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public Color? Color
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "srgbClr")
                return new Color(a: 255, r: 0, g: 0, b: 0);
            var val = el.Attribute("val")?.Value;
            if (val is null || val.Length < 6)
                return new Color(a: 255, r: 0, g: 0, b: 0);
            var r = Convert.ToByte(val[..2], 16);
            var g = Convert.ToByte(val[2..4], 16);
            var b = Convert.ToByte(val[4..6], 16);
            var a = ReadAlpha(el);
            return new Color(a: a, r: r, g: g, b: b);
        }
        set
        {
            if (_parentElement is null || value is null) return;
            ClearColorElements();
            var hex = $"{value.R:X2}{value.G:X2}{value.B:X2}";
            var srgb = new XElement(ANs + "srgbClr", new XAttribute("val", hex));
            _parentElement.Add(srgb);
            WriteAlpha(srgb, value.A);
            Save();
        }
    }

    /// <inheritdoc/>
    public PresetColor PresetColor
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "prstClr")
                return PresetColor.NotDefined;
            var val = el.Attribute("val")?.Value;
            if (val is null) return PresetColor.NotDefined;
            var pascalName = CamelToPascal(val);
            if (Enum.TryParse<PresetColor>(pascalName, out var result))
                return result;
            return PresetColor.NotDefined;
        }
        set
        {
            if (value == PresetColor.NotDefined) return;
            if (_parentElement is null) return;
            ClearColorElements();
            var camelName = PascalToCamel(value.ToString());
            _parentElement.Add(new XElement(ANs + "prstClr", new XAttribute("val", camelName)));
            Save();
        }
    }

    /// <inheritdoc/>
    public SchemeColor SchemeColor
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "schemeClr")
                return SchemeColor.NotDefined;
            var val = el.Attribute("val")?.Value;
            if (val is null) return SchemeColor.NotDefined;
            return SchemeClrToEnum.GetValueOrDefault(val, SchemeColor.NotDefined);
        }
        set
        {
            if (value == SchemeColor.NotDefined) return;
            if (_parentElement is null) return;
            if (!EnumToSchemeClr.TryGetValue(value, out var xmlVal)) return;
            ClearColorElements();
            _parentElement.Add(new XElement(ANs + "schemeClr", new XAttribute("val", xmlVal)));
            Save();
        }
    }

    /// <inheritdoc/>
    public int R
    {
        get => Color?.R ?? 0;
        set { }
    }

    /// <inheritdoc/>
    public int G
    {
        get => Color?.G ?? 0;
        set { }
    }

    /// <inheritdoc/>
    public int B
    {
        get => Color?.B ?? 0;
        set { }
    }

    /// <inheritdoc/>
    public float FloatR
    {
        get => R / 255f;
        set { }
    }

    /// <inheritdoc/>
    public float FloatG
    {
        get => G / 255f;
        set { }
    }

    /// <inheritdoc/>
    public float FloatB
    {
        get => B / 255f;
        set { }
    }

    /// <inheritdoc/>
    public float Hue
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return 0f;
            var val = el.Attribute("hue")?.Value;
            if (val is null || !int.TryParse(val, out var hue)) return 0f;
            return hue / 60000f;
        }
        set
        {
            if (_parentElement is null) return;
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return;
            el.SetAttributeValue("hue", (int)(value * 60000f));
            Save();
        }
    }

    /// <inheritdoc/>
    public float Saturation
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return 0f;
            var val = el.Attribute("sat")?.Value;
            if (val is null || !int.TryParse(val, out var sat)) return 0f;
            return sat / 1000f;
        }
        set
        {
            if (_parentElement is null) return;
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return;
            el.SetAttributeValue("sat", (int)(value * 1000f));
            Save();
        }
    }

    /// <inheritdoc/>
    public float Luminance
    {
        get
        {
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return 0f;
            var val = el.Attribute("lum")?.Value;
            if (val is null || !int.TryParse(val, out var lum)) return 0f;
            return lum / 1000f;
        }
        set
        {
            if (_parentElement is null) return;
            var el = FindColorElement();
            if (el is null || el.Name.LocalName != "hlsClr") return;
            el.SetAttributeValue("lum", (int)(value * 1000f));
            Save();
        }
    }
}
