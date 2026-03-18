using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents properties for lines filling.
/// </summary>
public sealed class LineFillFormat : ILineFillFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly HashSet<XName> FillTags =
    [
        ANs + "noFill",
        ANs + "solidFill",
        ANs + "gradFill",
        ANs + "pattFill",
    ];

    private static readonly HashSet<string> AfterFillLocals =
    [
        "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd", "extLst",
    ];

    private XElement? _lnElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes internal state.
    /// </summary>
    internal void InitInternal(XElement lnElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _lnElement = lnElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    private XElement? FindFillElement()
    {
        if (_lnElement is null) return null;
        foreach (var child in _lnElement.Elements())
        {
            if (FillTags.Contains(child.Name))
                return child;
        }
        return null;
    }

    private void RemoveFillElements()
    {
        if (_lnElement is null) return;
        foreach (var child in _lnElement.Elements().ToList())
        {
            if (FillTags.Contains(child.Name))
                child.Remove();
        }
    }

    private XElement InsertFillElement(XName tag)
    {
        var el = new XElement(tag);
        if (_lnElement is not null)
        {
            XElement? insertBefore = null;
            foreach (var child in _lnElement.Elements())
            {
                if (AfterFillLocals.Contains(child.Name.LocalName))
                {
                    insertBefore = child;
                    break;
                }
            }

            if (insertBefore is not null)
                insertBefore.AddBeforeSelf(el);
            else
                _lnElement.Add(el);
        }
        return el;
    }

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private void ResetStyleLnRef()
    {
        var spPr = _lnElement?.Parent;
        if (spPr is null) return;
        var shapeEl = spPr.Parent;
        if (shapeEl is null) return;
        var styleEl = shapeEl.Element(PNs + "style");
        if (styleEl is null) return;
        var lnRef = styleEl.Element(ANs + "lnRef");
        lnRef?.SetAttributeValue("idx", "0");
    }

    private void Save() => _slidePart?.Save();

    /// <inheritdoc/>
    public FillType FillType
    {
        get
        {
            var el = FindFillElement();
            if (el is null) return FillType.NotDefined;
            return el.Name.LocalName switch
            {
                "noFill" => FillType.NoFill,
                "solidFill" => FillType.Solid,
                "gradFill" => FillType.Gradient,
                "pattFill" => FillType.Pattern,
                _ => FillType.NotDefined,
            };
        }
        set
        {
            var tagLocal = value switch
            {
                FillType.NoFill => "noFill",
                FillType.Solid => "solidFill",
                FillType.Gradient => "gradFill",
                FillType.Pattern => "pattFill",
                _ => (string?)null,
            };

            var existing = FindFillElement();
            if (existing is not null && tagLocal is not null && existing.Name.LocalName == tagLocal)
                return;

            RemoveFillElements();
            if (tagLocal is not null)
            {
                InsertFillElement(ANs + tagLocal);
                ResetStyleLnRef();
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IColorFormat SolidFillColor
    {
        get
        {
            var el = FindFillElement();
            if (el is null || el.Name.LocalName != "solidFill")
            {
                RemoveFillElements();
                el = InsertFillElement(ANs + "solidFill");
            }
            var cf = new ColorFormat();
            cf.InitInternal(el, _parentSlide, _slidePart);
            return cf;
        }
    }

    /// <inheritdoc/>
    public IGradientFormat GradientFormat
    {
        get
        {
            var el = FindFillElement();
            if (el is null || el.Name.LocalName != "gradFill")
            {
                RemoveFillElements();
                el = InsertFillElement(ANs + "gradFill");
            }
            var gf = new GradientFormat();
            gf.InitInternal(el, _slidePart, _parentSlide);
            return gf;
        }
    }

    /// <inheritdoc/>
    public IPatternFormat PatternFormat
    {
        get
        {
            var el = FindFillElement();
            if (el is null || el.Name.LocalName != "pattFill")
            {
                RemoveFillElements();
                el = InsertFillElement(ANs + "pattFill");
            }
            var pf = new PatternFormat();
            pf.InitInternal(el, _parentSlide, _slidePart);
            return pf;
        }
    }

    /// <inheritdoc/>
    public NullableBool RotateWithShape
    {
        get
        {
            var el = FindFillElement();
            if (el is null) return NullableBool.NotDefined;
            var val = el.Attribute("rotWithShape")?.Value;
            return val switch
            {
                "1" => NullableBool.True,
                "0" => NullableBool.False,
                _ => NullableBool.NotDefined,
            };
        }
        set
        {
            var el = FindFillElement();
            if (el is null) return;
            if (value == NullableBool.NotDefined)
                el.Attribute("rotWithShape")?.Remove();
            else
                el.SetAttributeValue("rotWithShape", value == NullableBool.True ? "1" : "0");
            Save();
        }
    }
}
