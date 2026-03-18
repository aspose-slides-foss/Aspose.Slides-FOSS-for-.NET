using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a connector shape that can link two shapes via connection sites.
/// </summary>
public sealed class Connector : GeometryShape, IConnector
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private IBaseSlide? _connectorParentSlide;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal new void InitInternal(XElement? element, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        base.InitInternal(element, slidePart, parentSlide);
        _connectorParentSlide = parentSlide;
    }

    /// <inheritdoc/>
    public override bool IsTextHolder => false;

    /// <inheritdoc/>
    public override IPlaceholder? Placeholder => null;

    /// <inheritdoc/>
    public override ICustomData? CustomData => null;

    /// <inheritdoc/>
    public override IShapeStyle? ShapeStyle => null;

    /// <inheritdoc/>
    public override ShapeType ShapeType
    {
        get
        {
            if (_element is null)
                return ShapeType.NotDefined;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
                return ShapeType.NotDefined;

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
                return ShapeType.NotDefined;

            var prst = prstGeom.Attribute("prst")?.Value;
            return OoxmlPresetMapping.FromPreset(prst);
        }
        set
        {
            if (_element is null)
                return;
            if (value == ShapeType.NotDefined || value == ShapeType.Custom)
                return;

            var preset = OoxmlPresetMapping.ToPreset(value);
            if (preset is null)
                return;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
            {
                spPr = new XElement(PNs + "spPr");
                _element.Add(spPr);
            }

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
            {
                prstGeom = new XElement(ANs + "prstGeom");
                spPr.Add(prstGeom);
            }

            prstGeom.SetAttributeValue("prst", preset);

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public override IAdjustValueCollection? Adjustments
    {
        get
        {
            if (_element is null)
                return null;

            var collection = new AdjustValueCollection();

            var spPr = _element.Element(PNs + "spPr");
            var prstGeom = spPr?.Element(ANs + "prstGeom");
            var avLst = prstGeom?.Element(ANs + "avLst");
            if (avLst is not null)
                collection.InitInternal(avLst, _slidePart);

            return collection;
        }
    }

    /// <inheritdoc/>
    public IConnectorLock? ConnectorLock => null;

    /// <inheritdoc/>
    public IShape? StartShapeConnectedTo
    {
        get
        {
            var stCxn = GetConnectionElement("stCxn");
            if (stCxn is null)
                return null;

            var id = stCxn.Attribute("id")?.Value;
            if (id is null)
                return null;

            return FindShapeById(int.Parse(id));
        }
        set
        {
            if (value is null)
            {
                GetConnectionElement("stCxn")?.Remove();
                _slidePart?.Save();
                return;
            }

            SetConnectionElement("stCxn", value);
            _slidePart?.Save();
            Reroute();
        }
    }

    /// <inheritdoc/>
    public IShape? EndShapeConnectedTo
    {
        get
        {
            var endCxn = GetConnectionElement("endCxn");
            if (endCxn is null)
                return null;

            var id = endCxn.Attribute("id")?.Value;
            if (id is null)
                return null;

            return FindShapeById(int.Parse(id));
        }
        set
        {
            if (value is null)
            {
                GetConnectionElement("endCxn")?.Remove();
                _slidePart?.Save();
                return;
            }

            SetConnectionElement("endCxn", value);
            _slidePart?.Save();
            Reroute();
        }
    }

    /// <inheritdoc/>
    public int StartShapeConnectionSiteIndex
    {
        get
        {
            var stCxn = GetConnectionElement("stCxn");
            var idx = stCxn?.Attribute("idx")?.Value;
            return idx is not null ? int.Parse(idx) : 0;
        }
        set
        {
            var stCxn = GetConnectionElement("stCxn");
            stCxn?.SetAttributeValue("idx", value.ToString());
            _slidePart?.Save();
            Reroute();
        }
    }

    /// <inheritdoc/>
    public int EndShapeConnectionSiteIndex
    {
        get
        {
            var endCxn = GetConnectionElement("endCxn");
            var idx = endCxn?.Attribute("idx")?.Value;
            return idx is not null ? int.Parse(idx) : 0;
        }
        set
        {
            var endCxn = GetConnectionElement("endCxn");
            endCxn?.SetAttributeValue("idx", value.ToString());
            _slidePart?.Save();
            Reroute();
        }
    }

    /// <inheritdoc/>
    public IGeometryShape AsIGeometryShape => this;

    /// <inheritdoc/>
    public void Reroute()
    {
        var startShape = StartShapeConnectedTo;
        var endShape = EndShapeConnectedTo;
        if (startShape is null || endShape is null || _element is null)
            return;

        var (sx, sy) = GetConnectionPoint(startShape, StartShapeConnectionSiteIndex);
        var (ex, ey) = GetConnectionPoint(endShape, EndShapeConnectionSiteIndex);

        var xfrm = EnsureXfrm();

        var flipH = ex < sx;
        var flipV = ey < sy;

        var off = xfrm.Element(ANs + "off");
        if (off is null)
        {
            off = new XElement(ANs + "off");
            xfrm.Add(off);
        }

        var ext = xfrm.Element(ANs + "ext");
        if (ext is null)
        {
            ext = new XElement(ANs + "ext");
            xfrm.Add(ext);
        }

        off.SetAttributeValue("x", (int)MathF.Round(MathF.Min(sx, ex) * EmuPerPoint));
        off.SetAttributeValue("y", (int)MathF.Round(MathF.Min(sy, ey) * EmuPerPoint));
        ext.SetAttributeValue("cx", (int)MathF.Round(MathF.Abs(ex - sx) * EmuPerPoint));
        ext.SetAttributeValue("cy", (int)MathF.Round(MathF.Abs(ey - sy) * EmuPerPoint));

        if (flipH)
            xfrm.SetAttributeValue("flipH", "1");
        else
            xfrm.Attribute("flipH")?.Remove();

        if (flipV)
            xfrm.SetAttributeValue("flipV", "1");
        else
            xfrm.Attribute("flipV")?.Remove();

        _slidePart?.Save();
    }

    private XElement? GetCNvCxnSpPr()
    {
        if (_element is null)
            return null;

        var nvCxnSpPr = _element.Element(PNs + "nvCxnSpPr");
        return nvCxnSpPr?.Element(PNs + "cNvCxnSpPr");
    }

    private XElement EnsureCNvCxnSpPr()
    {
        if (_element is null)
            throw new InvalidOperationException("XML element is null.");

        var nvCxnSpPr = _element.Element(PNs + "nvCxnSpPr")
            ?? throw new InvalidOperationException("p:nvCxnSpPr element is missing.");

        var cNvCxnSpPr = nvCxnSpPr.Element(PNs + "cNvCxnSpPr");
        if (cNvCxnSpPr is null)
        {
            cNvCxnSpPr = new XElement(PNs + "cNvCxnSpPr");
            nvCxnSpPr.Add(cNvCxnSpPr);
        }

        return cNvCxnSpPr;
    }

    private XElement? GetConnectionElement(string localName)
    {
        var cNvCxnSpPr = GetCNvCxnSpPr();
        return cNvCxnSpPr?.Element(ANs + localName);
    }

    private void SetConnectionElement(string localName, IShape targetShape)
    {
        var cNvCxnSpPr = EnsureCNvCxnSpPr();

        // Get the target shape's cNvPr@id by navigating the target's XML.
        // We need to find the id from the shape - look for cNvPr in standard locations.
        var targetId = GetShapeCNvPrId(targetShape);

        var cxnElement = cNvCxnSpPr.Element(ANs + localName);
        if (cxnElement is null)
        {
            cxnElement = new XElement(ANs + localName);
            cNvCxnSpPr.Add(cxnElement);
        }

        cxnElement.SetAttributeValue("id", targetId);

        if (cxnElement.Attribute("idx") is null)
        {
            cxnElement.SetAttributeValue("idx", "0");
        }
    }

    private static string GetShapeCNvPrId(IShape shape)
    {
        if (shape is Shape s)
        {
            var cNvPr = s.GetCNvPr();
            var id = cNvPr?.Attribute("id")?.Value;
            if (id is not null)
                return id;
        }

        return "0";
    }

    /// <summary>
    /// Returns the connection point (x, y) in points for a connection site on a shape.
    /// Uses a 4-site model: 0 = top-center, 1 = left-center, 2 = bottom-center, 3 = right-center.
    /// Out-of-range index returns shape center.
    /// </summary>
    internal static (float X, float Y) GetConnectionPoint(IShape shape, int siteIndex)
    {
        var x = shape.X;
        var y = shape.Y;
        var w = shape.Width;
        var h = shape.Height;

        return siteIndex switch
        {
            0 => (x + w / 2f, y),
            1 => (x, y + h / 2f),
            2 => (x + w / 2f, y + h),
            3 => (x + w, y + h / 2f),
            _ => (x + w / 2f, y + h / 2f),
        };
    }

    private IShape? FindShapeById(int shapeId)
    {
        if (_connectorParentSlide is null)
            return null;

        var shapes = _connectorParentSlide.Shapes;
        if (shapes is null)
            return null;

        foreach (var shape in shapes)
        {
            if (shape is Shape s)
            {
                var cNvPr = s.GetCNvPr();
                if (cNvPr is not null)
                {
                    var idStr = cNvPr.Attribute("id")?.Value;
                    if (idStr is not null && int.Parse(idStr) == shapeId)
                        return shape;
                }
            }
        }

        return null;
    }
}
