using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Represents the presentation.xml part inside an OPC package.
/// Provides access to the parsed XML and relationship data.
/// </summary>
internal sealed class PresentationPart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string PartName = "ppt/presentation.xml";
    private const string RelsPartName = "ppt/_rels/presentation.xml.rels";

    private OpcPackage _package = null!;
    private XDocument? _doc;
    private RelsManager? _rels;

    /// <summary>
    /// Creates a <see cref="PresentationPart"/> from the given OPC package.
    /// </summary>
    internal static PresentationPart CreateFromPackage(OpcPackage package)
    {
        var part = new PresentationPart();
        part._package = package;
        return part;
    }

    /// <summary>
    /// Gets the parsed XML document for presentation.xml.
    /// </summary>
    internal XDocument Document
    {
        get
        {
            if (_doc is null)
            {
                var data = _package.GetPart(PartName);
                if (data is not null)
                {
                    using var ms = new MemoryStream(data);
                    _doc = XDocument.Load(ms);
                }
                else
                {
                    _doc = new XDocument(new XElement(PNs + "presentation"));
                }
            }
            return _doc;
        }
    }

    /// <summary>
    /// Gets the relationship manager for presentation.xml.rels.
    /// </summary>
    internal RelsManager Rels
    {
        get
        {
            if (_rels is null)
            {
                _rels = new RelsManager();
                var data = _package.GetPart(RelsPartName);
                if (data is not null)
                    _rels.Load(data);
            }
            return _rels;
        }
    }

    /// <summary>
    /// Gets the backing OPC package.
    /// </summary>
    internal OpcPackage Package => _package;

    /// <summary>
    /// Reads the first slide number from presentation.xml.
    /// Returns 1 if the attribute is not present.
    /// </summary>
    internal int ReadFirstSlideNumber()
    {
        var root = Document.Root;
        if (root is null)
            return 1;

        var attr = root.Attribute("firstSlideNum");
        if (attr is not null && int.TryParse(attr.Value, out var num))
            return num;

        return 1;
    }

    /// <summary>
    /// Writes the first slide number attribute to presentation.xml.
    /// </summary>
    internal void SetFirstSlideNumber(int number)
    {
        var root = Document.Root;
        root?.SetAttributeValue("firstSlideNum", number);
    }

    /// <summary>
    /// Gets the slide relationship entries from the sldIdLst element.
    /// Returns pairs of (slideId, relationshipId).
    /// </summary>
    internal List<(uint SlideId, string RelId)> GetSlideEntries()
    {
        var result = new List<(uint, string)>();
        var root = Document.Root;
        if (root is null)
            return result;

        var sldIdLst = root.Element(PNs + "sldIdLst");
        if (sldIdLst is null)
            return result;

        foreach (var sldId in sldIdLst.Elements(PNs + "sldId"))
        {
            var id = sldId.Attribute("id")?.Value;
            var rId = sldId.Attribute(RNs + "id")?.Value;
            if (id is not null && rId is not null && uint.TryParse(id, out var slideId))
                result.Add((slideId, rId));
        }

        return result;
    }

    /// <summary>
    /// Gets the slide master relationship entries from the sldMasterIdLst element.
    /// Returns pairs of (masterId, relationshipId).
    /// </summary>
    internal List<(uint MasterId, string RelId)> GetMasterEntries()
    {
        var result = new List<(uint, string)>();
        var root = Document.Root;
        if (root is null)
            return result;

        var masterIdLst = root.Element(PNs + "sldMasterIdLst");
        if (masterIdLst is null)
            return result;

        foreach (var masterId in masterIdLst.Elements(PNs + "sldMasterId"))
        {
            var id = masterId.Attribute("id")?.Value;
            var rId = masterId.Attribute(RNs + "id")?.Value;
            if (id is not null && rId is not null && uint.TryParse(id, out var mid))
                result.Add((mid, rId));
        }

        return result;
    }

    /// <summary>
    /// Flushes the in-memory XML back to the OPC package.
    /// </summary>
    internal void Flush()
    {
        if (_doc is not null)
        {
            using var ms = new MemoryStream();
            _doc.Save(ms);
            _package.SetPart(PartName, ms.ToArray());
        }

        if (_rels is not null)
        {
            _package.SetPart(RelsPartName, _rels.ToBytes());
        }
    }
}
