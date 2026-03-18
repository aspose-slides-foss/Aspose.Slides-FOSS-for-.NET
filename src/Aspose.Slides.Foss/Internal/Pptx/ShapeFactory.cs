using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// Factory for creating the correct <see cref="IShape"/> subclass from a PPTX XML element.
/// Inspects the element tag and, for graphic frames, the <c>graphicData</c> URI
/// to determine the concrete shape type.
/// </summary>
internal static class ShapeFactory
{
    private static readonly XNamespace ANs = Constants.Namespaces["a"];

    /// <summary>
    /// Creates an <see cref="IShape"/> wrapper for the given XML element.
    /// </summary>
    /// <param name="element">The XML element representing the shape (e.g., p:sp, p:pic).</param>
    /// <param name="slidePart">The <see cref="SlidePart"/> that owns this shape's XML.</param>
    /// <param name="parentSlide">The parent slide or layout that contains this shape.</param>
    /// <returns>An appropriate <see cref="IShape"/> instance, or <c>null</c> if the element type is not recognized.</returns>
    internal static IShape? CreateShape(XElement element, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        var localName = element.Name.LocalName;

        // For graphicFrame elements, delegate to CreateGraphicalObject which inspects the URI.
        if (localName == "graphicFrame")
        {
            return CreateGraphicalObject(element, slidePart, parentSlide);
        }

        Shape? shape = localName switch
        {
            "sp" => new AutoShape(),
            "pic" => new PictureFrame(),
            "grpSp" => new GroupShape(),
            "cxnSp" => new Connector(),
            _ => null,
        };

        if (shape is null)
            return null;

        shape.InitInternal(element, slidePart, parentSlide);
        return shape;
    }

    /// <summary>
    /// Creates the appropriate graphical object (Table, Chart, SmartArt, etc.) from a
    /// <c>p:graphicFrame</c> element by inspecting the <c>a:graphicData</c> URI.
    /// </summary>
    /// <param name="element">The <c>p:graphicFrame</c> XML element.</param>
    /// <param name="slidePart">The <see cref="SlidePart"/> that owns this shape's XML.</param>
    /// <param name="parentSlide">The parent slide or layout that contains this shape.</param>
    /// <returns>A <see cref="Table"/> or other graphical object, or <c>null</c> if the URI is not recognized.</returns>
    internal static IShape? CreateGraphicalObject(XElement element, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        // Find the a:graphicData element (may be nested under a:graphic)
        var graphicData = element.Descendants(ANs + "graphicData").FirstOrDefault();
        if (graphicData is null)
            return null;

        var uri = graphicData.Attribute("uri")?.Value ?? string.Empty;

        Shape? shape = null;

        if (uri.Contains("table", StringComparison.OrdinalIgnoreCase))
        {
            shape = new Table();
        }
        // Chart and SmartArt/Diagram types are not yet implemented.
        // When they are added, uncomment the corresponding branches:
        // else if (uri.Contains("chart", StringComparison.OrdinalIgnoreCase))
        // {
        //     shape = new Chart();
        // }
        // else if (uri.Contains("smartart", StringComparison.OrdinalIgnoreCase)
        //       || uri.Contains("diagram", StringComparison.OrdinalIgnoreCase))
        // {
        //     shape = new SmartArt();
        // }

        if (shape is null)
            return null;

        shape.InitInternal(element, slidePart, parentSlide);
        return shape;
    }
}
