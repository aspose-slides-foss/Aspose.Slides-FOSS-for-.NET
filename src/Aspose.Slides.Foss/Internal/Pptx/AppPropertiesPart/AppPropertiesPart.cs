using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.AppPropertiesPart;

/// <summary>
/// Parses and serializes docProps/app.xml (extended properties) in an OPC package.
/// Handles: Application, AppVersion, Company, Manager, PresentationFormat,
/// Template, TotalTime, Slides, HiddenSlides, Notes, Paragraphs, Words,
/// MMClips, ScaleCrop, LinksUpToDate, SharedDoc, HyperlinksChanged,
/// HyperlinkBase, HeadingPairs, TitlesOfParts.
/// </summary>
public sealed class AppPropertiesPart
{
    private static readonly XNamespace EpNs = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace VtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>
    /// The part path within the OPC package.
    /// </summary>
    public const string PartName = "docProps/app.xml";

    private readonly Internal.OpcPackage _package;
    private XElement? _root;
    private bool _dirty;

    // String properties
    /// <summary>Gets or sets the application name.</summary>
    public string? Application { get; set; }

    /// <summary>Gets or sets the application version.</summary>
    public string? AppVersion { get; set; }

    /// <summary>Gets or sets the company name.</summary>
    public string? Company { get; set; }

    /// <summary>Gets or sets the manager name.</summary>
    public string? Manager { get; set; }

    /// <summary>Gets or sets the presentation format.</summary>
    public string? PresentationFormat { get; set; }

    /// <summary>Gets or sets the template name.</summary>
    public string? Template { get; set; }

    /// <summary>Gets or sets the hyperlink base URI.</summary>
    public string? HyperlinkBase { get; set; }

    // Integer properties
    /// <summary>Gets or sets the total editing time in minutes.</summary>
    public int? TotalTime { get; set; }

    /// <summary>Gets or sets the number of slides.</summary>
    public int? Slides { get; set; }

    /// <summary>Gets or sets the number of hidden slides.</summary>
    public int? HiddenSlides { get; set; }

    /// <summary>Gets or sets the number of notes pages.</summary>
    public int? Notes { get; set; }

    /// <summary>Gets or sets the number of paragraphs.</summary>
    public int? Paragraphs { get; set; }

    /// <summary>Gets or sets the number of words.</summary>
    public int? Words { get; set; }

    /// <summary>Gets or sets the number of multimedia clips.</summary>
    public int? MmClips { get; set; }

    // Boolean properties
    /// <summary>Gets or sets whether the thumbnail is cropped to fit.</summary>
    public bool? ScaleCrop { get; set; }

    /// <summary>Gets or sets whether links are up to date.</summary>
    public bool? LinksUpToDate { get; set; }

    /// <summary>Gets or sets whether this is a shared document.</summary>
    public bool? SharedDoc { get; set; }

    /// <summary>Gets or sets whether hyperlinks have changed.</summary>
    public bool? HyperlinksChanged { get; set; }

    // Vector properties
    /// <summary>Gets the list of heading pairs.</summary>
    public List<HeadingPairData> HeadingPairs { get; set; } = [];

    /// <summary>Gets the list of part titles.</summary>
    public List<string> TitlesOfParts { get; set; } = [];

    /// <summary>
    /// Initializes the app properties part by parsing docProps/app.xml from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation.</param>
    internal AppPropertiesPart(Internal.OpcPackage package)
    {
        _package = package;
        Parse();
    }

    /// <summary>
    /// Parses the docProps/app.xml part from the package.
    /// </summary>
    public void Parse()
    {
        var data = _package.GetPart(PartName);
        if (data is null)
            return;

        using var stream = new MemoryStream(data);
        var doc = XDocument.Load(stream);
        _root = doc.Root;

        Application = GetText("Application");
        AppVersion = GetText("AppVersion");
        Company = GetText("Company");
        Manager = GetText("Manager");
        PresentationFormat = GetText("PresentationFormat");
        Template = GetText("Template");
        HyperlinkBase = GetText("HyperlinkBase");

        TotalTime = GetInt("TotalTime");
        Slides = GetInt("Slides");
        HiddenSlides = GetInt("HiddenSlides");
        Notes = GetInt("Notes");
        Paragraphs = GetInt("Paragraphs");
        Words = GetInt("Words");
        MmClips = GetInt("MMClips");

        ScaleCrop = GetBool("ScaleCrop");
        LinksUpToDate = GetBool("LinksUpToDate");
        SharedDoc = GetBool("SharedDoc");
        HyperlinksChanged = GetBool("HyperlinksChanged");

        ParseHeadingPairs();
        ParseTitlesOfParts();
    }

    /// <summary>
    /// Gets the text content of a named element from the root.
    /// </summary>
    /// <param name="localName">The local element name.</param>
    /// <returns>The text content, or <c>null</c> if not found or empty.</returns>
    public string? GetText(string localName)
    {
        if (_root is null)
            return null;
        var el = _root.Element(EpNs + localName);
        if (el is not null && !string.IsNullOrEmpty(el.Value))
            return el.Value;
        return null;
    }

    /// <summary>
    /// Gets the integer value of a named element from the root.
    /// </summary>
    /// <param name="localName">The local element name.</param>
    /// <returns>The parsed integer, or <c>null</c> if not found or not a valid integer.</returns>
    public int? GetInt(string localName)
    {
        var text = GetText(localName);
        if (text is not null && int.TryParse(text, out var value))
            return value;
        return null;
    }

    /// <summary>
    /// Gets the boolean value of a named element from the root.
    /// </summary>
    /// <param name="localName">The local element name.</param>
    /// <returns>The parsed boolean, or <c>null</c> if not found.</returns>
    public bool? GetBool(string localName)
    {
        var text = GetText(localName);
        if (text is not null)
            return text.Equals("true", StringComparison.OrdinalIgnoreCase) || text == "1";
        return null;
    }

    /// <summary>
    /// Parses HeadingPairs from the vt:vector element.
    /// Pairs come as alternating (name_variant, count_variant) entries.
    /// </summary>
    public void ParseHeadingPairs()
    {
        HeadingPairs.Clear();
        if (_root is null)
            return;

        var hpEl = _root.Element(EpNs + "HeadingPairs");
        if (hpEl is null)
            return;

        var vector = hpEl.Element(VtNs + "vector");
        if (vector is null)
            return;

        var variants = vector.Elements(VtNs + "variant").ToList();
        var i = 0;
        while (i + 1 < variants.Count)
        {
            var nameVariant = variants[i];
            var countVariant = variants[i + 1];

            var nameEl = nameVariant.Element(VtNs + "lpstr");
            var countEl = countVariant.Element(VtNs + "i4");

            if (nameEl is not null && countEl is not null)
            {
                var name = nameEl.Value ?? "";
                var count = int.TryParse(countEl.Value, out var c) ? c : 0;
                HeadingPairs.Add(new HeadingPairData(name, count));
            }

            i += 2;
        }
    }

    /// <summary>
    /// Parses TitlesOfParts from the vt:vector element.
    /// </summary>
    public void ParseTitlesOfParts()
    {
        TitlesOfParts.Clear();
        if (_root is null)
            return;

        var tpEl = _root.Element(EpNs + "TitlesOfParts");
        if (tpEl is null)
            return;

        var vector = tpEl.Element(VtNs + "vector");
        if (vector is null)
            return;

        foreach (var lpstr in vector.Elements(VtNs + "lpstr"))
        {
            TitlesOfParts.Add(lpstr.Value ?? "");
        }
    }

    /// <summary>
    /// Marks the part as dirty so it will be serialized on save.
    /// </summary>
    public void MarkDirty()
    {
        _dirty = true;
    }

    /// <summary>
    /// Serializes the app properties back to the OPC package.
    /// </summary>
    public void Save()
    {
        if (!_dirty && _root is not null)
            return;

        var root = new XElement(EpNs + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", VtNs));

        SetText(root, "Application", Application);
        SetText(root, "AppVersion", AppVersion);
        SetText(root, "Company", Company);
        SetText(root, "Manager", Manager);
        SetText(root, "PresentationFormat", PresentationFormat);
        SetText(root, "Template", Template);
        SetText(root, "HyperlinkBase", HyperlinkBase);

        SetInt(root, "TotalTime", TotalTime);
        SetInt(root, "Slides", Slides);
        SetInt(root, "HiddenSlides", HiddenSlides);
        SetInt(root, "Notes", Notes);
        SetInt(root, "Paragraphs", Paragraphs);
        SetInt(root, "Words", Words);
        SetInt(root, "MMClips", MmClips);

        SetBool(root, "ScaleCrop", ScaleCrop);
        SetBool(root, "LinksUpToDate", LinksUpToDate);
        SetBool(root, "SharedDoc", SharedDoc);
        SetBool(root, "HyperlinksChanged", HyperlinksChanged);

        if (HeadingPairs.Count > 0)
            WriteHeadingPairs(root);
        if (TitlesOfParts.Count > 0)
            WriteTitlesOfParts(root);

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());
        _dirty = false;
    }

    /// <summary>
    /// Sets a text element on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="value">The text value, or <c>null</c> to skip.</param>
    public void SetText(XElement root, string localName, string? value)
    {
        if (value is not null)
        {
            var el = new XElement(EpNs + localName) { Value = value };
            root.Add(el);
        }
    }

    /// <summary>
    /// Sets an integer element on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="value">The integer value, or <c>null</c> to skip.</param>
    public void SetInt(XElement root, string localName, int? value)
    {
        if (value is not null)
        {
            var el = new XElement(EpNs + localName) { Value = value.Value.ToString() };
            root.Add(el);
        }
    }

    /// <summary>
    /// Sets a boolean element on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="value">The boolean value, or <c>null</c> to skip.</param>
    public void SetBool(XElement root, string localName, bool? value)
    {
        if (value is not null)
        {
            var el = new XElement(EpNs + localName) { Value = value.Value ? "true" : "false" };
            root.Add(el);
        }
    }

    /// <summary>
    /// Writes heading pairs as a vt:vector of alternating name/count variants.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    public void WriteHeadingPairs(XElement root)
    {
        var hpEl = new XElement(EpNs + "HeadingPairs");
        var vector = new XElement(VtNs + "vector",
            new XAttribute("size", (HeadingPairs.Count * 2).ToString()),
            new XAttribute("baseType", "variant"));
        hpEl.Add(vector);

        foreach (var pair in HeadingPairs)
        {
            var v1 = new XElement(VtNs + "variant",
                new XElement(VtNs + "lpstr", pair.Name));
            vector.Add(v1);

            var v2 = new XElement(VtNs + "variant",
                new XElement(VtNs + "i4", pair.Count.ToString()));
            vector.Add(v2);
        }

        root.Add(hpEl);
    }

    /// <summary>
    /// Writes titles of parts as a vt:vector of lpstr elements.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    public void WriteTitlesOfParts(XElement root)
    {
        var tpEl = new XElement(EpNs + "TitlesOfParts");
        var vector = new XElement(VtNs + "vector",
            new XAttribute("size", TitlesOfParts.Count.ToString()),
            new XAttribute("baseType", "lpstr"));
        tpEl.Add(vector);

        foreach (var title in TitlesOfParts)
        {
            vector.Add(new XElement(VtNs + "lpstr", title));
        }

        root.Add(tpEl);
    }

    /// <summary>
    /// Resets all properties to null/empty.
    /// </summary>
    public void Clear()
    {
        Application = null;
        AppVersion = null;
        Company = null;
        Manager = null;
        PresentationFormat = null;
        Template = null;
        HyperlinkBase = null;
        TotalTime = null;
        Slides = null;
        HiddenSlides = null;
        Notes = null;
        Paragraphs = null;
        Words = null;
        MmClips = null;
        ScaleCrop = null;
        LinksUpToDate = null;
        SharedDoc = null;
        HyperlinksChanged = null;
        HeadingPairs = [];
        TitlesOfParts = [];
        _dirty = true;
    }
}
