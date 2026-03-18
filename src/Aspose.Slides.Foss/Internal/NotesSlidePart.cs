using System.Text;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages a notes slide XML part within the OPC package.
/// </summary>
internal sealed class NotesSlidePart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string NotesSlideRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

    private const string SlideRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    private readonly Dictionary<string, string> _placeholders = new();
    private XDocument? _document;
    private OpcPackage? _package;
    private string _partName = string.Empty;

    /// <summary>
    /// Gets the OPC part name for this notes slide.
    /// </summary>
    internal string PartName => _partName;

    /// <summary>
    /// Gets the notes slide relationship type constant.
    /// </summary>
    internal static string RelType => NotesSlideRelType;

    /// <summary>
    /// Determines whether a placeholder of the specified type exists.
    /// </summary>
    internal bool HasPlaceholder(string typeKey)
    {
        return _placeholders.ContainsKey(typeKey);
    }

    /// <summary>
    /// Adds a placeholder of the specified type if it does not already exist.
    /// </summary>
    internal void AddPlaceholder(string typeKey)
    {
        _placeholders.TryAdd(typeKey, string.Empty);
    }

    /// <summary>
    /// Removes a placeholder of the specified type if it exists.
    /// </summary>
    internal void RemovePlaceholder(string typeKey)
    {
        _placeholders.Remove(typeKey);
    }

    /// <summary>
    /// Sets the text content of a placeholder of the specified type.
    /// </summary>
    internal void SetPlaceholderText(string typeKey, string text)
    {
        _placeholders[typeKey] = text;
    }

    /// <summary>
    /// Gets the txBody element from the notes body placeholder shape.
    /// </summary>
    internal XElement? GetNotesTxBody()
    {
        var spTree = _document?.Root?.Element(PNs + "cSld")?.Element(PNs + "spTree");
        if (spTree is null) return null;

        foreach (var sp in spTree.Elements(PNs + "sp"))
        {
            var phType = sp.Element(PNs + "nvSpPr")
                ?.Element(PNs + "nvPr")
                ?.Element(PNs + "ph")
                ?.Attribute("type")?.Value;
            if (phType == "body")
                return sp.Element(PNs + "txBody");
        }

        return null;
    }

    /// <summary>
    /// Flushes the current state (XML + placeholders) to the OPC package.
    /// </summary>
    internal void Flush()
    {
        if (_document is null || _package is null || string.IsNullOrEmpty(_partName))
            return;

        SyncPlaceholdersToXml();

        using var ms = new MemoryStream();
        _document.Save(ms);
        _package.SetPart(_partName, ms.ToArray());
    }

    /// <summary>
    /// Creates an empty notes slide part in the OPC package for the given slide.
    /// </summary>
    internal static NotesSlidePart CreateEmpty(OpcPackage package, string slidePartName)
    {
        var notesPartName = ComputeNotesPartName(package, slidePartName);
        var doc = BuildEmptyNotesXml();

        var part = new NotesSlidePart
        {
            _document = doc,
            _package = package,
            _partName = notesPartName,
        };

        // Store notes XML in OPC
        using var ms = new MemoryStream();
        doc.Save(ms);
        package.SetPart(notesPartName, ms.ToArray());

        // Create notes slide rels pointing back to the parent slide
        var notesRels = new RelsManager();
        var relativeSlideTarget = ComputeRelativeTarget(notesPartName, slidePartName);
        notesRels.Add(SlideRelType, relativeSlideTarget);
        package.SetPart(GetRelsPath(notesPartName), notesRels.ToBytes());

        return part;
    }

    /// <summary>
    /// Loads a NotesSlidePart from the OPC package.
    /// </summary>
    internal static NotesSlidePart? Load(OpcPackage package, string partName)
    {
        var data = package.GetPart(partName);
        if (data is null) return null;

        using var ms = new MemoryStream(data);
        var doc = XDocument.Load(ms);
        var part = new NotesSlidePart
        {
            _document = doc,
            _package = package,
            _partName = partName,
        };

        part.ParsePlaceholdersFromXml();
        return part;
    }

    /// <summary>
    /// Deletes a notes slide part and its rels from the OPC package.
    /// </summary>
    internal static void Delete(OpcPackage package, string partName)
    {
        package.RemovePart(partName);
        package.RemovePart(GetRelsPath(partName));
    }

    /// <summary>
    /// Computes the relative target path from one part to another.
    /// </summary>
    internal static string ComputeRelativeTarget(string fromPartName, string toPartName)
    {
        var fromSegments = GetDirectoryPath(fromPartName).TrimEnd('/').Split('/');
        var toDir = GetDirectoryPath(toPartName);
        var toFile = toPartName[toDir.Length..];

        int common = 0;
        var toSegments = toDir.TrimEnd('/').Split('/');
        while (common < fromSegments.Length && common < toSegments.Length &&
               string.Equals(fromSegments[common], toSegments[common], StringComparison.OrdinalIgnoreCase))
        {
            common++;
        }

        var sb = new StringBuilder();
        for (int i = common; i < fromSegments.Length; i++)
            sb.Append("../");
        for (int i = common; i < toSegments.Length; i++)
            sb.Append(toSegments[i]).Append('/');
        sb.Append(toFile);
        return sb.ToString();
    }

    // ── Private helpers ──────────────────────────────────────

    private void SyncPlaceholdersToXml()
    {
        var spTree = _document?.Root?.Element(PNs + "cSld")?.Element(PNs + "spTree");
        if (spTree is null) return;

        string[] types = ["ftr", "dt", "hdr", "sldNum"];
        int nextId = 10;

        foreach (var type in types)
        {
            var existingSp = FindPlaceholderShape(spTree, type);
            if (_placeholders.TryGetValue(type, out var text))
            {
                if (existingSp is null)
                {
                    spTree.Add(BuildPlaceholderShape(type, text, nextId++));
                }
                else if (text.Length > 0)
                {
                    SetShapeText(existingSp, text);
                }
            }
            else
            {
                existingSp?.Remove();
            }
        }
    }

    private void ParsePlaceholdersFromXml()
    {
        var spTree = _document?.Root?.Element(PNs + "cSld")?.Element(PNs + "spTree");
        if (spTree is null) return;

        foreach (var sp in spTree.Elements(PNs + "sp"))
        {
            var phType = sp.Element(PNs + "nvSpPr")
                ?.Element(PNs + "nvPr")
                ?.Element(PNs + "ph")
                ?.Attribute("type")?.Value;
            if (phType is "ftr" or "dt" or "hdr" or "sldNum")
            {
                _placeholders[phType] = GetShapeText(sp);
            }
        }
    }

    private static XElement? FindPlaceholderShape(XElement spTree, string phType)
    {
        foreach (var sp in spTree.Elements(PNs + "sp"))
        {
            var type = sp.Element(PNs + "nvSpPr")
                ?.Element(PNs + "nvPr")
                ?.Element(PNs + "ph")
                ?.Attribute("type")?.Value;
            if (type == phType)
                return sp;
        }
        return null;
    }

    private static XElement BuildPlaceholderShape(string phType, string text, int id)
    {
        return new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", id),
                    new XAttribute("name", $"Placeholder {id}")),
                new XElement(PNs + "cNvSpPr",
                    new XElement(ANs + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(PNs + "nvPr",
                    new XElement(PNs + "ph", new XAttribute("type", phType)))),
            new XElement(PNs + "spPr"),
            new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    string.IsNullOrEmpty(text)
                        ? (object)new XElement(ANs + "endParaRPr")
                        : new XElement(ANs + "r", new XElement(ANs + "t", text)))));
    }

    private static void SetShapeText(XElement sp, string text)
    {
        var txBody = sp.Element(PNs + "txBody");
        if (txBody is null) return;
        txBody.Elements(ANs + "p").Remove();
        txBody.Add(new XElement(ANs + "p",
            new XElement(ANs + "r", new XElement(ANs + "t", text))));
    }

    private static string GetShapeText(XElement sp)
    {
        var txBody = sp.Element(PNs + "txBody");
        if (txBody is null) return "";
        return string.Join("", txBody.Descendants(ANs + "t").Select(t => t.Value));
    }

    private static string ComputeNotesPartName(OpcPackage package, string slidePartName)
    {
        var slideFile = slidePartName[(slidePartName.LastIndexOf('/') + 1)..];
        var notesFile = slideFile.Replace("slide", "notesSlide", StringComparison.OrdinalIgnoreCase);
        var notesPartName = $"ppt/notesSlides/{notesFile}";

        int index = 1;
        while (package.GetPart(notesPartName) is not null)
        {
            index++;
            notesPartName = $"ppt/notesSlides/notesSlide{index}.xml";
        }

        return notesPartName;
    }

    private static XDocument BuildEmptyNotesXml()
    {
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(PNs + "notes",
                new XAttribute(XNamespace.Xmlns + "a", ANs),
                new XAttribute(XNamespace.Xmlns + "r", RNs),
                new XElement(PNs + "cSld",
                    new XElement(PNs + "spTree",
                        new XElement(PNs + "nvGrpSpPr",
                            new XElement(PNs + "cNvPr",
                                new XAttribute("id", "1"),
                                new XAttribute("name", "")),
                            new XElement(PNs + "cNvGrpSpPr"),
                            new XElement(PNs + "nvPr")),
                        new XElement(PNs + "grpSpPr"),
                        new XElement(PNs + "sp",
                            new XElement(PNs + "nvSpPr",
                                new XElement(PNs + "cNvPr",
                                    new XAttribute("id", "2"),
                                    new XAttribute("name", "Slide Image Placeholder 1")),
                                new XElement(PNs + "cNvSpPr",
                                    new XElement(ANs + "spLocks",
                                        new XAttribute("noGrp", "1"),
                                        new XAttribute("noRot", "1"),
                                        new XAttribute("noChangeAspect", "1"))),
                                new XElement(PNs + "nvPr",
                                    new XElement(PNs + "ph", new XAttribute("type", "sldImg")))),
                            new XElement(PNs + "spPr")),
                        new XElement(PNs + "sp",
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
                                    new XElement(ANs + "r",
                                        new XElement(ANs + "t", "")))))))));
    }

    private static string GetDirectoryPath(string partName)
    {
        var lastSlash = partName.LastIndexOf('/');
        return lastSlash >= 0 ? partName[..(lastSlash + 1)] : "";
    }

    private static string GetRelsPath(string partName)
    {
        var dir = GetDirectoryPath(partName);
        var file = partName[dir.Length..];
        return $"{dir}_rels/{file}.rels";
    }
}
