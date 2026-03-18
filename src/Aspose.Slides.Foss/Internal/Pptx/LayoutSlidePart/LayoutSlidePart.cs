using System.Collections.Frozen;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.LayoutSlidePart;

/// <summary>
/// Manages a layout slide XML part (ppt/slideLayouts/slideLayoutN.xml).
/// Read-only for now; provides access to layout properties.
/// </summary>
public sealed class LayoutSlidePart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string SlideMasterRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    /// <summary>
    /// Maps OOXML layout type attribute values to SlideLayoutType enum value strings.
    /// </summary>
    private static readonly FrozenDictionary<string, string> LayoutTypeMap = new Dictionary<string, string>
    {
        ["blank"] = "Blank",
        ["chart"] = "Chart",
        ["chartAndTx"] = "ChartAndText",
        ["clipArtAndTx"] = "ClipArtAndText",
        ["clipArtAndVertTx"] = "ClipArtAndVerticalText",
        ["cust"] = "Custom",
        ["dgm"] = "Diagram",
        ["fourObj"] = "FourObjects",
        ["mediaAndTx"] = "MediaAndText",
        ["obj"] = "Object",
        ["objAndTx"] = "ObjectAndText",
        ["objAndTwoObj"] = "ObjectAndTwoObject",
        ["objOnly"] = "ObjectOnly",
        ["objOverTx"] = "ObjectOverText",
        ["objTx"] = "ObjectText",
        ["picTx"] = "PictureAndCaption",
        ["secHead"] = "SectionHeader",
        ["tbl"] = "Table",
        ["title"] = "Title",
        ["titleOnly"] = "TitleOnly",
        ["twoColTx"] = "TwoColumnText",
        ["twoObj"] = "TwoObjects",
        ["twoObjAndObj"] = "TwoObjectsAndObject",
        ["twoObjAndTx"] = "TwoObjectsAndText",
        ["twoObjOverTx"] = "TwoObjectsOverText",
        ["twoTxTwoObj"] = "TwoTextAndTwoObjects",
        ["tx"] = "Text",
        ["txAndChart"] = "TextAndChart",
        ["txAndClipArt"] = "TextAndClipArt",
        ["txAndMedia"] = "TextAndMedia",
        ["txAndObj"] = "TextAndObject",
        ["txAndTwoObj"] = "TextAndTwoObjects",
        ["txOverObj"] = "TextOverObject",
        ["vertTitleAndTx"] = "VerticalTitleAndText",
        ["vertTitleAndTxOverChart"] = "VerticalTitleAndTextOverChart",
        ["vertTx"] = "VerticalText",
    }.ToFrozenDictionary();

    private readonly OpcPackage _package;
    private readonly string _partName;
    private XElement _root = null!;

    /// <summary>
    /// Initializes a layout slide part and loads its XML from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the layout.</param>
    /// <param name="partName">The part path (e.g., "ppt/slideLayouts/slideLayout1.xml").</param>
    internal LayoutSlidePart(OpcPackage package, string partName)
    {
        _package = package;
        _partName = partName;
        Load();
    }

    /// <summary>
    /// Loads and parses the layout slide XML from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(_partName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            var doc = XDocument.Load(ms);
            _root = doc.Root ?? throw new InvalidOperationException("Layout slide XML has no root element");
        }
        else
        {
            throw new InvalidOperationException($"Layout slide part not found: {_partName}");
        }
    }

    /// <summary>
    /// Gets the part name of this layout slide.
    /// </summary>
    public string PartName => _partName;

    /// <summary>
    /// Gets or sets the layout name from the p:cSld element's name attribute.
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
    /// Gets the raw layout type string from the p:sldLayout type attribute.
    /// Defaults to "cust" if not specified.
    /// </summary>
    public string LayoutTypeRaw => (string?)_root.Attribute("type") ?? "cust";

    /// <summary>
    /// Gets the SlideLayoutType enum value string for this layout.
    /// Defaults to "Custom" for unrecognized types.
    /// </summary>
    public string LayoutTypeValue
    {
        get
        {
            var raw = LayoutTypeRaw;
            return LayoutTypeMap.GetValueOrDefault(raw, "Custom");
        }
    }

    /// <summary>
    /// Gets the master slide part name resolved from this layout's relationships.
    /// Returns <c>null</c> if no slide master relationship is found.
    /// </summary>
    public string? MasterPartName
    {
        get
        {
            var relsPartName = GetRelsPartName(_partName);
            var relsContent = _package.GetPart(relsPartName);
            if (relsContent is null)
                return null;

            using var msRels = new MemoryStream(relsContent);
            var doc = XDocument.Load(msRels);
            var rel = doc.Root?.Elements(RelsNs + "Relationship")
                .FirstOrDefault(e => (string?)e.Attribute("Type") == SlideMasterRelType);

            if (rel is null)
                return null;

            var target = (string?)rel.Attribute("Target") ?? "";
            return ResolveTarget(target);
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
    /// Saves the layout slide XML back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(_partName, ms.ToArray());
    }

    private static string GetRelsPartName(string sourcePart)
    {
        if (string.IsNullOrEmpty(sourcePart))
            return "_rels/.rels";

        var lastSlash = sourcePart.LastIndexOf('/');
        if (lastSlash >= 0)
            return $"{sourcePart[..lastSlash]}/_rels/{sourcePart[(lastSlash + 1)..]}.rels";

        return $"_rels/{sourcePart}.rels";
    }
}
