using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Base class for all shapes on a slide.
/// </summary>
public class Shape : PVIObject, IShape
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace AdecNs = "http://schemas.microsoft.com/office/drawing/2017/decorative";

    /// <summary>EMUs per point (1 point = 12700 EMU).</summary>
    internal const float EmuPerPoint = 12700f;

    /// <summary>Rotation unit: 60000ths of a degree per degree.</summary>
    private const float RotationUnit = 60000f;

    /// <summary>The backing XML element for this shape.</summary>
    private protected XElement? _element;

    /// <summary>
    /// Returns the backing XML element for this shape.
    /// Used internally for element identity comparisons.
    /// </summary>
    internal XElement? GetElement() => _element;

    /// <summary>The slide part containing the slide XML.</summary>
    private protected SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="element">The XML element representing this shape.</param>
    /// <param name="slidePart">The slide part containing the slide XML.</param>
    /// <param name="parentSlide">The parent slide object.</param>
    internal virtual void InitInternal(XElement? element, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _element = element;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <inheritdoc/>
    public virtual bool IsTextHolder => false;

    /// <inheritdoc/>
    public virtual IPlaceholder? Placeholder => null;

    /// <inheritdoc/>
    public virtual ICustomData? CustomData { get; }

    /// <inheritdoc/>
    public virtual ILineFormat LineFormat
    {
        get
        {
            if (_element is null)
                throw new InvalidOperationException("Shape is not initialized.");

            var spPr = EnsureSpPr();
            var lf = new LineFormat();
            lf.InitInternal(spPr, _parentSlide);
            return lf;
        }
    }

    /// <inheritdoc/>
    public virtual IThreeDFormat ThreeDFormat
    {
        get
        {
            if (_element is null)
                throw new InvalidOperationException("Shape is not initialized.");

            var spPr = EnsureSpPr();
            var tdf = new ThreeDFormat();
            tdf.InitInternal(spPr, _parentSlide);
            return tdf;
        }
    }

    /// <inheritdoc/>
    public virtual IEffectFormat EffectFormat
    {
        get
        {
            if (_element is null)
                throw new InvalidOperationException("Shape is not initialized.");

            var spPr = EnsureSpPr();
            var ef = new EffectFormat();
            ef.InitInternal(spPr, _parentSlide);
            return ef;
        }
    }

    /// <inheritdoc/>
    public virtual IFillFormat FillFormat
    {
        get
        {
            if (_element is null)
                throw new InvalidOperationException("Shape is not initialized.");

            var spPr = EnsureSpPr();
            var ff = new FillFormat();
            ff.InitInternal(spPr, _parentSlide);
            return ff;
        }
    }

    /// <inheritdoc/>
    public virtual IShapeFrame RawFrame
    {
        get => BuildFrame();
        set => ApplyFrame(value);
    }

    /// <inheritdoc/>
    public virtual IShapeFrame Frame
    {
        get => BuildFrame();
        set => ApplyFrame(value);
    }

    /// <inheritdoc/>
    public virtual bool Hidden
    {
        get
        {
            var cNvPr = GetCNvPr();
            return cNvPr?.Attribute("hidden")?.Value == "1";
        }
        set
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return;

            if (value)
                cNvPr.SetAttributeValue("hidden", "1");
            else
                cNvPr.Attribute("hidden")?.Remove();

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual int ZOrderPosition
    {
        get
        {
            if (_element?.Parent is null)
                return 0;

            var index = 0;
            foreach (var child in _element.Parent.Elements())
            {
                if (child == _element)
                    return index;
                index++;
            }

            return 0;
        }
    }

    /// <inheritdoc/>
    public virtual int ConnectionSiteCount => 0;

    /// <inheritdoc/>
    public virtual float Rotation
    {
        get
        {
            var xfrm = GetXfrm();
            var rot = xfrm?.Attribute("rot")?.Value;
            if (rot is null)
                return 0f;

            return int.Parse(rot) / RotationUnit;
        }
        set
        {
            var xfrm = EnsureXfrm();
            xfrm.SetAttributeValue("rot", (int)MathF.Round(value * RotationUnit));
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual float X
    {
        get => GetOffsetValue("x");
        set => SetOffsetValue("x", value);
    }

    /// <inheritdoc/>
    public virtual float Y
    {
        get => GetOffsetValue("y");
        set => SetOffsetValue("y", value);
    }

    /// <inheritdoc/>
    public virtual float Width
    {
        get => GetExtentValue("cx");
        set => SetExtentValue("cx", value);
    }

    /// <inheritdoc/>
    public virtual float Height
    {
        get => GetExtentValue("cy");
        set => SetExtentValue("cy", value);
    }

    /// <inheritdoc/>
    public virtual int UniqueId
    {
        get
        {
            var cNvPr = GetCNvPr();
            var id = cNvPr?.Attribute("id")?.Value;
            return id is not null ? int.Parse(id) : 0;
        }
    }

    /// <inheritdoc/>
    public virtual int OfficeInteropShapeId
    {
        get
        {
            var cNvPr = GetCNvPr();
            var id = cNvPr?.Attribute("id")?.Value;
            return id is not null ? int.Parse(id) : 0;
        }
    }

    /// <inheritdoc/>
    public virtual string AlternativeText
    {
        get => GetCNvPr()?.Attribute("descr")?.Value ?? string.Empty;
        set
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return;

            cNvPr.SetAttributeValue("descr", value);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual string AlternativeTextTitle
    {
        get => GetCNvPr()?.Attribute("title")?.Value ?? string.Empty;
        set
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return;

            cNvPr.SetAttributeValue("title", value);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual string Name
    {
        get => GetCNvPr()?.Attribute("name")?.Value ?? string.Empty;
        set
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return;

            cNvPr.SetAttributeValue("name", value);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual bool IsDecorative
    {
        get
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return false;

            var extLst = cNvPr.Element(ANs + "extLst");
            if (extLst is null)
                return false;

            foreach (var ext in extLst.Elements(ANs + "ext"))
            {
                var decorative = ext.Element(AdecNs + "decorative");
                if (decorative is not null)
                    return decorative.Attribute("val")?.Value == "1";
            }

            return false;
        }
        set
        {
            var cNvPr = GetCNvPr();
            if (cNvPr is null)
                return;

            var extLst = cNvPr.Element(ANs + "extLst");

            if (value)
            {
                if (extLst is null)
                {
                    extLst = new XElement(ANs + "extLst");
                    cNvPr.Add(extLst);
                }

                // Check if decorative extension already exists
                foreach (var ext in extLst.Elements(ANs + "ext"))
                {
                    var existing = ext.Element(AdecNs + "decorative");
                    if (existing is not null)
                    {
                        existing.SetAttributeValue("val", "1");
                        _slidePart?.Save();
                        return;
                    }
                }

                // Add new extension element
                var newExt = new XElement(ANs + "ext",
                    new XAttribute("uri", "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"),
                    new XElement(AdecNs + "decorative",
                        new XAttribute(XNamespace.Xmlns + "adec", AdecNs.NamespaceName),
                        new XAttribute("val", "1")));
                extLst.Add(newExt);
            }
            else
            {
                if (extLst is not null)
                {
                    foreach (var ext in extLst.Elements(ANs + "ext").ToList())
                    {
                        var decorative = ext.Element(AdecNs + "decorative");
                        if (decorative is not null)
                        {
                            ext.Remove();
                            break;
                        }
                    }
                }
            }

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public virtual bool IsGrouped
    {
        get
        {
            if (_element?.Parent is null)
                return false;

            return _element.Parent.Name == PNs + "grpSp";
        }
    }

    /// <inheritdoc/>
    public virtual ISlideComponent AsISlideComponent => this;

    // ------- Helper methods -------

    /// <summary>
    /// Finds the shape properties element (spPr or grpSpPr).
    /// </summary>
    internal XElement? GetSpPr()
    {
        if (_element is null)
            return null;

        return _element.Element(PNs + "spPr")
            ?? _element.Element(PNs + "grpSpPr");
    }

    /// <summary>
    /// Returns existing spPr or creates a new one.
    /// </summary>
    internal XElement EnsureSpPr()
    {
        if (_element is null)
            throw new InvalidOperationException("XML element is not set.");

        var spPr = _element.Element(PNs + "spPr");
        if (spPr is not null)
            return spPr;

        var grpSpPr = _element.Element(PNs + "grpSpPr");
        if (grpSpPr is not null)
            return grpSpPr;

        spPr = new XElement(PNs + "spPr");
        _element.Add(spPr);
        return spPr;
    }

    /// <summary>
    /// Finds the cNvPr element from the shape XML by checking all nvXxxPr container variants.
    /// </summary>
    internal XElement? GetCNvPr()
    {
        if (_element is null)
            return null;

        ReadOnlySpan<string> nvNames =
        [
            "nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"
        ];

        foreach (var nvName in nvNames)
        {
            var nvElem = _element.Element(PNs + nvName);
            if (nvElem is not null)
                return nvElem.Element(PNs + "cNvPr");
        }

        return null;
    }

    /// <summary>
    /// Finds the <c>a:xfrm</c> element directly within a shape XML element.
    /// Checks spPr, grpSpPr, and p:xfrm (for graphic frames).
    /// </summary>
    internal static XElement? FindXfrmInElement(XElement xmlElement)
    {
        // Try spPr (most common: sp, pic, cxnSp)
        var spPr = xmlElement.Element(PNs + "spPr");
        if (spPr is not null)
        {
            var xfrm = spPr.Element(ANs + "xfrm");
            if (xfrm is not null)
                return xfrm;
        }

        // Try grpSpPr (group shapes)
        var grpSpPr = xmlElement.Element(PNs + "grpSpPr");
        if (grpSpPr is not null)
        {
            var xfrm = grpSpPr.Element(ANs + "xfrm");
            if (xfrm is not null)
                return xfrm;
        }

        // Try p:xfrm (graphic frames like tables, charts)
        return xmlElement.Element(PNs + "xfrm");
    }

    /// <summary>
    /// Gets placeholder type and index from a <c>&lt;p:ph&gt;</c> element if this shape is a placeholder.
    /// </summary>
    /// <returns>A tuple of (type, idx) or <c>null</c> if not a placeholder.</returns>
    internal (string? Type, string Idx)? GetPlaceholderInfo()
    {
        if (_element is null)
            return null;

        ReadOnlySpan<string> nvNames =
        [
            "nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"
        ];

        foreach (var nvName in nvNames)
        {
            var nvElem = _element.Element(PNs + nvName);
            if (nvElem is not null)
            {
                var nvPr = nvElem.Element(PNs + "nvPr");
                if (nvPr is not null)
                {
                    var ph = nvPr.Element(PNs + "ph");
                    if (ph is not null)
                        return (ph.Attribute("type")?.Value, ph.Attribute("idx")?.Value ?? "0");
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Finds the xfrm for a matching placeholder in a layout or master XML root.
    /// </summary>
    internal static XElement? FindPlaceholderXfrmInXml(XElement root, string? phType, string phIdx)
    {
        var spTree = root.Descendants(PNs + "spTree").FirstOrDefault();
        if (spTree is null)
            return null;

        ReadOnlySpan<string> nvNames =
        [
            "nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"
        ];

        foreach (var child in spTree.Elements())
        {
            XElement? ph = null;
            foreach (var nvName in nvNames)
            {
                var nvElem = child.Element(PNs + nvName);
                if (nvElem is not null)
                {
                    var nvPr = nvElem.Element(PNs + "nvPr");
                    if (nvPr is not null)
                    {
                        ph = nvPr.Element(PNs + "ph");
                        break;
                    }
                }
            }

            if (ph is null)
                continue;

            var childType = ph.Attribute("type")?.Value;
            var childIdx = ph.Attribute("idx")?.Value ?? "0";

            if (childType == phType && childIdx == phIdx)
            {
                var xfrm = FindXfrmInElement(child);
                if (xfrm is not null)
                    return xfrm;
            }
        }

        return null;
    }

    /// <summary>
    /// Walks the layout to master chain to find inherited xfrm for placeholder shapes.
    /// </summary>
    internal XElement? GetInheritedXfrm()
    {
        var phInfo = GetPlaceholderInfo();
        if (phInfo is null || _slidePart is null)
            return null;

        var (phType, phIdx) = phInfo.Value;
        var package = _slidePart.Package;
        if (package is null)
            return null;

        var layoutPartName = _slidePart.LayoutPartName;
        if (layoutPartName is null)
            return null;

        var layoutContent = package.GetPart(layoutPartName);
        if (layoutContent is null)
            return null;

        using var msLayout = new MemoryStream(layoutContent);
        var layoutRoot = XDocument.Load(msLayout).Root;
        if (layoutRoot is null)
            return null;

        var xfrm = FindPlaceholderXfrmInXml(layoutRoot, phType, phIdx);
        if (xfrm is not null)
            return xfrm;

        // Try master slide (resolve from layout's relationships)
        var layoutPart = new Aspose.Slides.Foss.Internal.Pptx.LayoutSlidePart.LayoutSlidePart(package, layoutPartName);
        var masterPartName = layoutPart.MasterPartName;
        if (masterPartName is not null)
        {
            var masterContent = package.GetPart(masterPartName);
            if (masterContent is not null)
            {
                using var msMaster = new MemoryStream(masterContent);
                var masterRoot = XDocument.Load(msMaster).Root;
                if (masterRoot is not null)
                {
                    xfrm = FindPlaceholderXfrmInXml(masterRoot, phType, phIdx);
                    if (xfrm is not null)
                        return xfrm;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the <c>a:xfrm</c> element from the shape XML.
    /// For placeholder shapes with no local xfrm, walks the layout to master inheritance chain.
    /// </summary>
    internal XElement? GetXfrm()
    {
        if (_element is null)
            return null;

        var xfrm = FindXfrmInElement(_element);
        if (xfrm is not null)
            return xfrm;

        // Placeholder inheritance: try layout, then master
        return GetInheritedXfrm();
    }

    /// <summary>
    /// Gets or creates the <c>a:xfrm</c> element.
    /// </summary>
    internal XElement EnsureXfrm()
    {
        var xfrm = GetXfrm();
        if (xfrm is not null)
            return xfrm;

        if (_element is null)
            throw new InvalidOperationException("Shape has no XML element");

        // Create xfrm under the appropriate parent
        var spPr = _element.Element(PNs + "spPr");
        if (spPr is null)
        {
            var grpSpPr = _element.Element(PNs + "grpSpPr");
            if (grpSpPr is not null)
                spPr = grpSpPr;
            else
            {
                spPr = new XElement(PNs + "spPr");
                _element.Add(spPr);
            }
        }

        xfrm = new XElement(ANs + "xfrm",
            new XElement(ANs + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
            new XElement(ANs + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")));
        spPr.Add(xfrm);
        return xfrm;
    }

    /// <summary>
    /// Builds a <see cref="ShapeFrame"/> from the current xfrm element.
    /// </summary>
    internal IShapeFrame BuildFrame()
    {
        var xfrm = GetXfrm();
        if (xfrm is null)
            return new ShapeFrame(0, 0, 0, 0, NullableBool.NotDefined, NullableBool.NotDefined, 0);

        var off = xfrm.Element(ANs + "off");
        var ext = xfrm.Element(ANs + "ext");

        var x = ParseEmu(off?.Attribute("x")?.Value);
        var y = ParseEmu(off?.Attribute("y")?.Value);
        var cx = ParseEmu(ext?.Attribute("cx")?.Value);
        var cy = ParseEmu(ext?.Attribute("cy")?.Value);

        var rot = xfrm.Attribute("rot")?.Value;
        var rotation = rot is not null ? int.Parse(rot) / RotationUnit : 0f;

        var flipH = xfrm.Attribute("flipH")?.Value == "1" ? NullableBool.True : NullableBool.False;
        var flipV = xfrm.Attribute("flipV")?.Value == "1" ? NullableBool.True : NullableBool.False;

        return new ShapeFrame(x, y, cx, cy, flipH, flipV, rotation);
    }

    /// <summary>
    /// Writes a <see cref="IShapeFrame"/> back to the xfrm element.
    /// </summary>
    internal void ApplyFrame(IShapeFrame value)
    {
        var xfrm = EnsureXfrm();

        var off = xfrm.Element(ANs + "off");
        if (off is null)
        {
            off = new XElement(ANs + "off", new XAttribute("x", "0"), new XAttribute("y", "0"));
            xfrm.Add(off);
        }

        var ext = xfrm.Element(ANs + "ext");
        if (ext is null)
        {
            ext = new XElement(ANs + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0"));
            xfrm.Add(ext);
        }

        off.SetAttributeValue("x", (int)MathF.Round(value.X * EmuPerPoint));
        off.SetAttributeValue("y", (int)MathF.Round(value.Y * EmuPerPoint));
        ext.SetAttributeValue("cx", (int)MathF.Round(value.Width * EmuPerPoint));
        ext.SetAttributeValue("cy", (int)MathF.Round(value.Height * EmuPerPoint));

        xfrm.SetAttributeValue("rot", (int)MathF.Round(value.Rotation * RotationUnit));

        if (value.FlipH == NullableBool.True)
            xfrm.SetAttributeValue("flipH", "1");
        else
            xfrm.Attribute("flipH")?.Remove();

        if (value.FlipV == NullableBool.True)
            xfrm.SetAttributeValue("flipV", "1");
        else
            xfrm.Attribute("flipV")?.Remove();

        _slidePart?.Save();
    }

    /// <summary>
    /// Reads an EMU value from an a:off attribute and converts to points.
    /// </summary>
    private float GetOffsetValue(string attr)
    {
        var xfrm = GetXfrm();
        var off = xfrm?.Element(ANs + "off");
        return ParseEmu(off?.Attribute(attr)?.Value);
    }

    /// <summary>
    /// Writes a points value to an a:off attribute as EMU.
    /// </summary>
    private void SetOffsetValue(string attr, float value)
    {
        var xfrm = EnsureXfrm();

        var off = xfrm.Element(ANs + "off");
        if (off is null)
        {
            off = new XElement(ANs + "off");
            xfrm.Add(off);
        }

        off.SetAttributeValue(attr, (int)MathF.Round(value * EmuPerPoint));
        _slidePart?.Save();
    }

    /// <summary>
    /// Reads an EMU value from an a:ext attribute and converts to points.
    /// </summary>
    private float GetExtentValue(string attr)
    {
        var xfrm = GetXfrm();
        var ext = xfrm?.Element(ANs + "ext");
        return ParseEmu(ext?.Attribute(attr)?.Value);
    }

    /// <summary>
    /// Writes a points value to an a:ext attribute as EMU.
    /// </summary>
    private void SetExtentValue(string attr, float value)
    {
        var xfrm = EnsureXfrm();

        var ext = xfrm.Element(ANs + "ext");
        if (ext is null)
        {
            ext = new XElement(ANs + "ext");
            xfrm.Add(ext);
        }

        ext.SetAttributeValue(attr, (int)MathF.Round(value * EmuPerPoint));
        _slidePart?.Save();
    }

    /// <summary>
    /// Parses an EMU string value and converts to points.
    /// </summary>
    private static float ParseEmu(string? value)
    {
        if (value is null)
            return 0f;

        return long.Parse(value) / EmuPerPoint;
    }
}
