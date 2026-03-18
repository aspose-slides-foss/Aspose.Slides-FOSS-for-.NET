using System.Xml.Linq;
using Aspose.Slides.Foss.Internal.Opc.ContentTypes;
using Aspose.Slides.Foss.Internal.Opc.Relationships;

namespace Aspose.Slides.Foss.Internal.Pptx.NotesSlidePart;

/// <summary>
/// Manages a notes slide XML part (ppt/notesSlides/notesSlideN.xml).
/// Provides access to notes text and placeholder management.
/// </summary>
public sealed class NotesSlidePart
{
    private static readonly XNamespace PNs = Constants.Namespaces["p"];
    private static readonly XNamespace ANs = Constants.Namespaces["a"];

    /// <summary>Placeholder type for the notes body.</summary>
    private const string PhTypeBody = "body";

    /// <summary>Placeholder type for date/time.</summary>
    private const string PhTypeDt = "dt";

    /// <summary>Placeholder type for footer.</summary>
    private const string PhTypeFtr = "ftr";

    /// <summary>Placeholder type for header.</summary>
    private const string PhTypeHdr = "hdr";

    /// <summary>Placeholder type for slide number.</summary>
    private const string PhTypeSldNum = "sldNum";

    /// <summary>Placeholder type for slide image.</summary>
    private const string PhTypeSldImg = "sldImg";

    /// <summary>Text-bearing placeholder types that get an empty txBody when added.</summary>
    private static readonly HashSet<string> TextPhTypes = [PhTypeDt, PhTypeFtr, PhTypeHdr];

    private readonly Opc.OpcPackage _package;
    private readonly string _partName;
    private readonly RelationshipsManager _relsManager;
    private XElement _root = null!;

    /// <summary>
    /// Initializes a notes slide part from an existing part in the package.
    /// </summary>
    /// <param name="package">The OPC package containing the notes slide.</param>
    /// <param name="partName">The part path (e.g., "ppt/notesSlides/notesSlide1.xml").</param>
    internal NotesSlidePart(Opc.OpcPackage package, string partName)
    {
        _package = package;
        _partName = partName;
        _relsManager = new RelationshipsManager(package, partName);
        Load();
    }

    /// <summary>
    /// Gets the part name of this notes slide.
    /// </summary>
    public string PartName => _partName;

    /// <summary>
    /// Gets or sets the slide name from the p:cSld element's name attribute.
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
    /// Loads and parses the notes slide XML from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(_partName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            var doc = XDocument.Load(ms);
            _root = doc.Root ?? throw new InvalidOperationException("Notes slide XML has no root element");
        }
        else
        {
            throw new InvalidOperationException($"Notes slide part not found: {_partName}");
        }
    }

    /// <summary>
    /// Gets the spTree element from the notes slide.
    /// </summary>
    /// <returns>The spTree element, or <c>null</c> if not found.</returns>
    public XElement? GetSpTree()
    {
        return _root.Descendants(XName.Get("spTree", PNs.NamespaceName)).FirstOrDefault();
    }

    /// <summary>
    /// Finds the first placeholder shape with the given type.
    /// </summary>
    /// <param name="phType">The placeholder type string (e.g., "body", "ftr", "hdr").</param>
    /// <returns>The p:sp element, or <c>null</c> if not found.</returns>
    public XElement? FindPlaceholder(string phType)
    {
        var spTree = GetSpTree();
        if (spTree is null)
            return null;

        foreach (var sp in spTree.Elements(PNs + "sp"))
        {
            var nvPr = sp.Element(PNs + "nvSpPr")?.Element(PNs + "nvPr");
            if (nvPr is null)
                continue;

            var ph = nvPr.Element(PNs + "ph");
            if (ph is null)
                continue;

            if ((string?)ph.Attribute("type") == phType)
                return sp;
        }

        return null;
    }

    /// <summary>
    /// Gets the txBody element of the notes body placeholder.
    /// </summary>
    /// <returns>The txBody element of the body placeholder, or <c>null</c>.</returns>
    public XElement? GetNotesTxbody()
    {
        var bodySp = FindPlaceholder(PhTypeBody);
        return bodySp?.Element(PNs + "txBody");
    }

    /// <summary>
    /// Checks if a placeholder of the given type exists in the notes slide.
    /// </summary>
    /// <param name="phType">The placeholder type string.</param>
    /// <returns><c>true</c> if the placeholder exists; otherwise, <c>false</c>.</returns>
    public bool HasPlaceholder(string phType)
    {
        return FindPlaceholder(phType) is not null;
    }

    /// <summary>
    /// Removes the placeholder shape of the given type.
    /// </summary>
    /// <param name="phType">The placeholder type string to remove.</param>
    public void RemovePlaceholder(string phType)
    {
        var sp = FindPlaceholder(phType);
        sp?.Remove();
    }

    /// <summary>
    /// Adds a minimal placeholder shape of the given type if not already present.
    /// </summary>
    /// <param name="phType">The placeholder type string to add.</param>
    public void AddPlaceholder(string phType)
    {
        if (HasPlaceholder(phType))
            return;

        var spTree = GetSpTree();
        if (spTree is null)
            return;

        // Find the next available shape ID
        var maxId = 1;
        foreach (var sp in spTree.Elements(PNs + "sp"))
        {
            var cNvPr = sp.Element(PNs + "nvSpPr")?.Element(PNs + "cNvPr");
            if (cNvPr is not null && int.TryParse((string?)cNvPr.Attribute("id"), out var spId))
            {
                maxId = Math.Max(maxId, spId);
            }
        }

        var shapeId = maxId + 1;
        var spElem = BuildPlaceholderShape(phType, shapeId);
        spTree.Add(spElem);
    }

    /// <summary>
    /// Builds a minimal placeholder shape element for notes slides.
    /// </summary>
    /// <param name="phType">The placeholder type string.</param>
    /// <param name="shapeId">The shape ID to assign.</param>
    /// <returns>A new p:sp element.</returns>
    public static XElement BuildPlaceholderShape(string phType, int shapeId)
    {
        var sp = new XElement(PNs + "sp");

        var nvSpPr = new XElement(PNs + "nvSpPr");
        sp.Add(nvSpPr);

        var cNvPr = new XElement(PNs + "cNvPr",
            new XAttribute("id", shapeId.ToString()),
            new XAttribute("name", $"{phType} Placeholder {shapeId}"));
        nvSpPr.Add(cNvPr);

        var cNvSpPr = new XElement(PNs + "cNvSpPr",
            new XElement(ANs + "spLocks", new XAttribute("noGrp", "1")));
        nvSpPr.Add(cNvSpPr);

        var nvPr = new XElement(PNs + "nvPr",
            new XElement(PNs + "ph", new XAttribute("type", phType)));
        nvSpPr.Add(nvPr);

        // Empty shape properties
        sp.Add(new XElement(PNs + "spPr"));

        // Add empty text body for text-bearing placeholders
        if (TextPhTypes.Contains(phType))
        {
            var txBody = new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "endParaRPr")));
            sp.Add(txBody);
        }

        return sp;
    }

    /// <summary>
    /// Sets text content for a placeholder shape. Adds the placeholder if it does not exist.
    /// </summary>
    /// <param name="phType">The placeholder type string.</param>
    /// <param name="text">The text to set.</param>
    public void SetPlaceholderText(string phType, string text)
    {
        if (!HasPlaceholder(phType))
            AddPlaceholder(phType);

        var sp = FindPlaceholder(phType);
        if (sp is null)
            return;

        var txBody = sp.Element(PNs + "txBody");
        if (txBody is null)
        {
            txBody = new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"));
            sp.Add(txBody);
        }

        // Remove existing paragraphs
        txBody.Elements(ANs + "p").Remove();

        // Add new paragraph with the given text
        XElement aP;
        if (!string.IsNullOrEmpty(text))
        {
            aP = new XElement(ANs + "p",
                new XElement(ANs + "r",
                    new XElement(ANs + "t", text)));
        }
        else
        {
            aP = new XElement(ANs + "p",
                new XElement(ANs + "endParaRPr"));
        }

        txBody.Add(aP);
    }

    /// <summary>
    /// Saves the notes slide XML back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(_partName, ms.ToArray());
        _relsManager.Save();
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
    /// Creates a new empty notes slide in the package for a given slide.
    /// </summary>
    /// <param name="package">The OPC package.</param>
    /// <param name="slidePartName">The part name of the owning slide.</param>
    /// <returns>The newly created <see cref="NotesSlidePart"/>.</returns>
    internal static NotesSlidePart CreateEmpty(Opc.OpcPackage package, string slidePartName)
    {
        // Find the next available notes slide number
        var nextNum = 1;
        string partName;
        while (true)
        {
            var candidate = $"ppt/notesSlides/notesSlide{nextNum}.xml";
            if (!package.HasPart(candidate))
            {
                partName = candidate;
                break;
            }

            nextNum++;
        }

        // Build and store the notes slide XML
        var notesXml = BuildNotesXml();
        package.SetPart(partName, notesXml);

        // Create the notes slide's own relationships
        var relsManager = new RelationshipsManager(package, partName);

        // Relationship: notes slide -> parent slide
        var slideRelative = ComputeRelativeTarget(partName, slidePartName);
        relsManager.AddRelationship(RelConstants.RelTypes["slide"], slideRelative);

        // Relationship: notes slide -> notes master (if present)
        var notesMasterPartName = FindNotesMaster(package);
        if (notesMasterPartName is not null)
        {
            var masterRelative = ComputeRelativeTarget(partName, notesMasterPartName);
            relsManager.AddRelationship(RelConstants.RelTypes["notes_master"], masterRelative);
        }

        relsManager.Save();

        // Register content type
        var ctManager = new Opc.ContentTypes.ContentTypesManager(package);
        ctManager.AddOverride(partName, ContentTypeConstants.ContentTypes["notes_slide"]);
        ctManager.Save();

        return new NotesSlidePart(package, partName);
    }

    /// <summary>
    /// Finds the notes master part name in the package.
    /// </summary>
    /// <param name="package">The OPC package.</param>
    /// <returns>The part name of the notes master, or <c>null</c> if not found.</returns>
    internal static string? FindNotesMaster(Opc.OpcPackage package)
    {
        foreach (var part in package.GetPartNames())
        {
            if (part.StartsWith("ppt/notesMasters/", StringComparison.Ordinal) && part.EndsWith(".xml", StringComparison.Ordinal))
                return part;
        }

        return null;
    }

    /// <summary>
    /// Computes a relative path from one part to another.
    /// </summary>
    /// <param name="fromPart">The source part name.</param>
    /// <param name="toPart">The destination part name.</param>
    /// <returns>A relative path string.</returns>
    public static string ComputeRelativeTarget(string fromPart, string toPart)
    {
        var fromDir = fromPart.Contains('/')
            ? fromPart[..fromPart.LastIndexOf('/')]
            : "";
        var toDir = toPart.Contains('/')
            ? toPart[..toPart.LastIndexOf('/')]
            : "";
        var toFile = toPart[(toPart.LastIndexOf('/') + 1)..];

        if (fromDir == toDir)
            return toFile;

        var fromParts = string.IsNullOrEmpty(fromDir) ? [] : fromDir.Split('/');
        var toParts = string.IsNullOrEmpty(toDir) ? [] : toDir.Split('/');

        // Find common prefix length
        var commonLen = 0;
        for (var i = 0; i < Math.Min(fromParts.Length, toParts.Length); i++)
        {
            if (fromParts[i] == toParts[i])
                commonLen = i + 1;
            else
                break;
        }

        var upCount = fromParts.Length - commonLen;
        var downPath = string.Join('/', toParts[commonLen..]);

        var result = string.Concat(Enumerable.Repeat("../", upCount));
        if (!string.IsNullOrEmpty(downPath))
            result += downPath + "/";
        result += toFile;

        return result;
    }

    /// <summary>
    /// Builds a minimal notes slide XML with a slide image and body placeholder.
    /// </summary>
    /// <returns>UTF-8 encoded XML bytes.</returns>
    public static byte[] BuildNotesXml()
    {
        var notes = new XElement(PNs + "notes",
            new XAttribute(XNamespace.Xmlns + "a", ANs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", Constants.Namespaces["r"]),
            new XAttribute(XNamespace.Xmlns + "p", PNs.NamespaceName));

        // <p:cSld>
        var cSld = new XElement(PNs + "cSld");
        notes.Add(cSld);

        var spTree = new XElement(PNs + "spTree");
        cSld.Add(spTree);

        // Group shape header (required)
        var nvGrpSpPr = new XElement(PNs + "nvGrpSpPr",
            new XElement(PNs + "cNvPr",
                new XAttribute("id", "1"),
                new XAttribute("name", "")),
            new XElement(PNs + "cNvGrpSpPr"),
            new XElement(PNs + "nvPr"));
        spTree.Add(nvGrpSpPr);

        var grpSpPr = new XElement(PNs + "grpSpPr",
            new XElement(ANs + "xfrm",
                new XElement(ANs + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(ANs + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
                new XElement(ANs + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(ANs + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0"))));
        spTree.Add(grpSpPr);

        // Slide image placeholder (type="sldImg")
        var sp1 = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", "2"),
                    new XAttribute("name", "Slide Image Placeholder 1")),
                new XElement(PNs + "cNvSpPr",
                    new XElement(ANs + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(PNs + "nvPr",
                    new XElement(PNs + "ph", new XAttribute("type", "sldImg")))),
            new XElement(PNs + "spPr"));
        spTree.Add(sp1);

        // Notes body placeholder (type="body", idx="1")
        var sp2 = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", "3"),
                    new XAttribute("name", "Notes Placeholder 2")),
                new XElement(PNs + "cNvSpPr",
                    new XElement(ANs + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(PNs + "nvPr",
                    new XElement(PNs + "ph",
                        new XAttribute("type", "body"),
                        new XAttribute("idx", "1")))),
            new XElement(PNs + "spPr"),
            new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "endParaRPr"))));
        spTree.Add(sp2);

        // Color map override
        var clrMapOvr = new XElement(PNs + "clrMapOvr",
            new XElement(ANs + "masterClrMapping"));
        notes.Add(clrMapOvr);

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), notes);
        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Deletes a notes slide and its associated files from the package.
    /// </summary>
    /// <param name="package">The OPC package.</param>
    /// <param name="partName">The notes slide part name to delete.</param>
    internal static void Delete(Opc.OpcPackage package, string partName)
    {
        package.DeletePart(partName);

        var relsPartName = RelationshipsManager.GetRelsPartName(partName);
        package.DeletePart(relsPartName);

        var ctManager = new Opc.ContentTypes.ContentTypesManager(package);
        ctManager.RemoveOverride(partName);
        ctManager.Save();
    }
}
