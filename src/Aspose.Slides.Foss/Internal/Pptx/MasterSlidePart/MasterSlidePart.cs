using System.Xml.Linq;
using Aspose.Slides.Foss.Internal.Opc.Relationships;

namespace Aspose.Slides.Foss.Internal.Pptx.MasterSlidePart;

/// <summary>
/// Manages a master slide XML part (ppt/slideMasters/slideMasterN.xml).
/// Read-only for now; provides access to master slide properties.
/// </summary>
public sealed class MasterSlidePart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private readonly Opc.OpcPackage _package;
    private readonly string _partName;
    private readonly RelationshipsManager _relsManager;
    private XElement _root = null!;

    /// <summary>
    /// Initializes a master slide part and loads its XML from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the master slide.</param>
    /// <param name="partName">The part path (e.g., "ppt/slideMasters/slideMaster1.xml").</param>
    internal MasterSlidePart(Opc.OpcPackage package, string partName)
    {
        _package = package;
        _partName = partName;
        _relsManager = new RelationshipsManager(package, partName);
        Load();
    }

    /// <summary>
    /// Loads and parses the master slide XML from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(_partName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            var doc = XDocument.Load(ms);
            _root = doc.Root ?? throw new InvalidOperationException("Master slide XML has no root element");
        }
        else
        {
            throw new InvalidOperationException($"Master slide part not found: {_partName}");
        }
    }

    /// <summary>
    /// Gets the part name of this master slide.
    /// </summary>
    public string PartName => _partName;

    /// <summary>
    /// Gets or sets the master slide name from the p:cSld element's name attribute.
    /// </summary>
    public string Name
    {
        get
        {
            var csld = _root.Descendants(PNs + "cSld").FirstOrDefault();
            return (string?)csld?.Attribute("name") ?? "";
        }
        set
        {
            var csld = _root.Descendants(PNs + "cSld").FirstOrDefault();
            csld?.SetAttributeValue("name", value);
        }
    }

    /// <summary>
    /// Gets the list of layout slide part names from this master's relationships.
    /// </summary>
    public List<string> LayoutPartNames
    {
        get
        {
            var rels = _relsManager.GetRelationshipsByType(RelConstants.RelTypes["slide_layout"]);
            var result = new List<string>();
            foreach (var rel in rels)
            {
                result.Add(ResolveTarget(rel.Target));
            }
            return result;
        }
    }

    /// <summary>
    /// Resolves a relative target path to an absolute part name within the package.
    /// </summary>
    /// <param name="target">The target path, potentially relative.</param>
    /// <returns>The resolved absolute part name.</returns>
    public string ResolveTarget(string target)
    {
        if (target.StartsWith('/'))
            return target.TrimStart('/');

        var baseDir = _partName.Contains('/')
            ? _partName[.._partName.LastIndexOf('/')]
            : "";

        var segments = (baseDir + "/" + target).Split('/');
        var resolved = new List<string>();
        foreach (var segment in segments)
        {
            if (segment == "..")
            {
                if (resolved.Count > 0)
                    resolved.RemoveAt(resolved.Count - 1);
            }
            else if (segment is not ("" or "."))
            {
                resolved.Add(segment);
            }
        }
        return string.Join('/', resolved);
    }

    /// <summary>
    /// Saves the master slide XML back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(_partName, ms.ToArray());
        _relsManager.Save();
    }
}
