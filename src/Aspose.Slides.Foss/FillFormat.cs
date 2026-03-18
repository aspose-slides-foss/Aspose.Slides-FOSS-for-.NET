using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents fill formatting options.
/// </summary>
public sealed class FillFormat : PVIObject, IFillFormat, IFillParamSource
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly HashSet<XName> FillTags =
    [
        ANs + "noFill",
        ANs + "solidFill",
        ANs + "gradFill",
        ANs + "blipFill",
        ANs + "pattFill",
        ANs + "grpFill",
    ];

    /// <summary>
    /// Element locals that must appear after fill children in spPr ordering.
    /// </summary>
    private static readonly HashSet<string> AfterFillLocals =
    [
        "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst",
    ];

    private XElement? _parentElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state for this fill format.
    /// </summary>
    internal void InitInternal(XElement parentElement, IBaseSlide? parentSlide, SlidePart? slidePart = null)
    {
        _parentElement = parentElement;
        _parentSlide = parentSlide;
        _slidePart = slidePart;
    }

    internal XElement? FindFillElement()
    {
        if (_parentElement is null) return null;
        foreach (var child in _parentElement.Elements())
        {
            if (FillTags.Contains(child.Name))
                return child;
        }
        return null;
    }

    internal void RemoveFillElements()
    {
        if (_parentElement is null) return;
        foreach (var child in _parentElement.Elements().ToList())
        {
            if (FillTags.Contains(child.Name))
                child.Remove();
        }
    }

    internal XElement InsertFillElement(XName tag)
    {
        var el = new XElement(tag);
        if (_parentElement is not null)
        {
            XElement? insertBefore = null;
            foreach (var child in _parentElement.Elements())
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
                _parentElement.Add(el);
        }
        return el;
    }

    internal XElement GetOrCreateFill(string localName)
    {
        var el = FindFillElement();
        if (el is not null && el.Name.LocalName == localName)
            return el;
        RemoveFillElements();
        el = InsertFillElement(ANs + localName);
        AddFillDefaults(el);
        return el;
    }

    private static void AddFillDefaults(XElement el)
    {
        switch (el.Name.LocalName)
        {
            case "gradFill":
                el.Add(new XElement(ANs + "lin", new XAttribute("ang", "0"), new XAttribute("scaled", "1")));
                break;
            case "blipFill":
                el.Add(new XElement(ANs + "blip"));
                var stretch = new XElement(ANs + "stretch");
                stretch.Add(new XElement(ANs + "fillRect"));
                el.Add(stretch);
                break;
        }
    }

    internal void Save() => _slidePart?.Save();

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
                "blipFill" => FillType.Picture,
                "grpFill" => FillType.Group,
                _ => FillType.NotDefined,
            };
        }
        set
        {
            if (_parentElement is null) return;
            var tagLocal = value switch
            {
                FillType.NoFill => "noFill",
                FillType.Solid => "solidFill",
                FillType.Gradient => "gradFill",
                FillType.Pattern => "pattFill",
                FillType.Picture => "blipFill",
                FillType.Group => "grpFill",
                _ => (string?)null,
            };

            // Preserve existing fill element if type already matches
            var existing = FindFillElement();
            if (existing is not null && tagLocal is not null && existing.Name.LocalName == tagLocal)
                return;

            RemoveFillElements();
            if (tagLocal is not null)
            {
                var el = InsertFillElement(ANs + tagLocal);
                AddFillDefaults(el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IColorFormat SolidFillColor
    {
        get
        {
            var solidFill = GetOrCreateFill("solidFill");
            var cf = new ColorFormat();
            cf.InitInternal(solidFill, _parentSlide, _slidePart);
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
                el = GetOrCreateFill("gradFill");
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
                el = GetOrCreateFill("pattFill");
            var pf = new PatternFormat();
            pf.InitInternal(el, _parentSlide, _slidePart);
            return pf;
        }
    }

    /// <inheritdoc/>
    public IPictureFillFormat PictureFillFormat
    {
        get
        {
            var el = FindFillElement();
            if (el is null || el.Name.LocalName != "blipFill")
                el = GetOrCreateFill("blipFill");
            var pff = new PictureFillFormat();
            pff.InitInternal(el, _slidePart);
            return pff;
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
                "1" or "true" => NullableBool.True,
                "0" or "false" => NullableBool.False,
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
