using System.Xml;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// The XML namespaces an OOXML presentation package is written in.
/// </summary>
internal static class Ns
{
    /// <summary>PresentationML.</summary>
    internal static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    /// <summary>DrawingML.</summary>
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>Relationship references used inside part markup (<c>r:id</c>, <c>r:embed</c>, <c>r:link</c>).</summary>
    internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The Open Packaging Conventions relationships part.</summary>
    internal static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>The Open Packaging Conventions content-types part.</summary>
    internal static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>The PowerPoint 2010 extension namespace, which carries sections.</summary>
    internal static readonly XNamespace P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    /// <summary>Extended document properties (<c>docProps/app.xml</c>).</summary>
    internal static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    /// <summary>The relationship type of an embedded image.</summary>
    internal const string ImageRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>The relationship type of a slide.</summary>
    internal const string SlideRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    /// <summary>The relationship type of a notes master.</summary>
    internal const string NotesMasterRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    /// <summary>
    /// Builds a namespace manager carrying the prefixes used by the XPath expressions in these tests.
    /// </summary>
    internal static XmlNamespaceManager Resolver()
    {
        var manager = new XmlNamespaceManager(new NameTable());
        manager.AddNamespace("p", P.NamespaceName);
        manager.AddNamespace("a", A.NamespaceName);
        manager.AddNamespace("r", R.NamespaceName);
        manager.AddNamespace("rel", Rel.NamespaceName);
        manager.AddNamespace("ct", Ct.NamespaceName);
        manager.AddNamespace("p14", P14.NamespaceName);
        manager.AddNamespace("ep", Ep.NamespaceName);
        return manager;
    }
}
