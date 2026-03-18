using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the base class for shapes that have geometric properties.
/// </summary>
public abstract class GeometryShape : Shape, IGeometryShape
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <inheritdoc/>
    public virtual IShapeStyle? ShapeStyle => null;

    /// <inheritdoc/>
    public virtual ShapeType ShapeType
    {
        get => ShapeType.NotDefined;
        set { }
    }

    /// <inheritdoc/>
    public virtual IAdjustValueCollection? Adjustments
    {
        get
        {
            var coll = new AdjustValueCollection();
            var spPr = _element?.Element(PNs + "spPr");
            var prstGeom = spPr?.Element(ANs + "prstGeom");
            var avLst = prstGeom?.Element(ANs + "avLst");
            if (avLst is not null)
                coll.InitInternal(avLst, _slidePart);
            return coll;
        }
    }
}
