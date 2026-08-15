using System.Globalization;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages the OPC extended/app properties part (docProps/app.xml).
/// </summary>
internal sealed class AppPropertiesPart
{
    private static readonly XNamespace EpNs = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace VtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private static readonly string PartPath = "docProps/app.xml";

    private readonly XElement _root;
    private bool _isDirty;

    internal bool IsDirty => _isDirty;

    internal AppPropertiesPart(XElement root)
    {
        _root = root;
    }

    internal static AppPropertiesPart CreateFromPackage(OpcPackage package)
    {
        var data = package.GetPart(PartPath);
        if (data is not null)
        {
            using var stream = new MemoryStream(data);
            var doc = XDocument.Load(stream);
            return new AppPropertiesPart(doc.Root ?? throw new InvalidOperationException("App properties XML has no root element"));
        }

        var root = new XElement(EpNs + "Properties",
            new XAttribute(XNamespace.Xmlns + "ep", EpNs),
            new XAttribute(XNamespace.Xmlns + "vt", VtNs));
        return new AppPropertiesPart(root);
    }

    internal void MarkDirty() => _isDirty = true;

    internal string? AppVersion => GetElement("AppVersion");

    internal string? Application
    {
        get => GetElement("Application");
        set => SetElement("Application", value);
    }

    internal string? Company
    {
        get => GetElement("Company");
        set => SetElement("Company", value);
    }

    internal string? Manager
    {
        get => GetElement("Manager");
        set => SetElement("Manager", value);
    }

    internal string? PresentationFormat
    {
        get => GetElement("PresentationFormat") ?? GetElement("PresentatFormat");
        set
        {
            // Earlier versions wrote "PresentatFormat", which CT_Properties does not declare.
            _root.Element(EpNs + "PresentatFormat")?.Remove();
            SetElement("PresentationFormat", value);
        }
    }

    internal string? Template
    {
        get => GetElement("Template");
        set => SetElement("Template", value);
    }

    internal string? HyperlinkBase
    {
        get => GetElement("HyperlinkBase");
        set => SetElement("HyperlinkBase", value);
    }

    internal TimeSpan TotalTime
    {
        get
        {
            var text = GetElement("TotalTime");
            if (string.IsNullOrEmpty(text))
                return TimeSpan.Zero;
            return int.TryParse(text, CultureInfo.InvariantCulture, out var minutes)
                ? TimeSpan.FromMinutes(minutes)
                : TimeSpan.Zero;
        }
        set => SetElement("TotalTime", ((int)value.TotalMinutes).ToString(CultureInfo.InvariantCulture));
    }

    internal bool SharedDoc
    {
        get => GetBoolElement("SharedDoc");
        set => SetBoolElement("SharedDoc", value);
    }

    internal bool ScaleCrop
    {
        get => GetBoolElement("ScaleCrop");
        set => SetBoolElement("ScaleCrop", value);
    }

    internal bool LinksUpToDate
    {
        get => GetBoolElement("LinksUpToDate");
        set => SetBoolElement("LinksUpToDate", value);
    }

    internal bool HyperlinksChanged
    {
        get => GetBoolElement("HyperlinksChanged");
        set => SetBoolElement("HyperlinksChanged", value);
    }

    internal int Slides => GetIntElement("Slides");
    internal int HiddenSlides => GetIntElement("HiddenSlides");
    internal int Notes => GetIntElement("Notes");
    internal int Paragraphs => GetIntElement("Paragraphs");
    internal int Words => GetIntElement("Words");
    internal int MultimediaClips => GetIntElement("MMClips");

    internal IReadOnlyList<HeadingPairData> HeadingPairs
    {
        get
        {
            var result = new List<HeadingPairData>();
            var hpElem = _root.Element(EpNs + "HeadingPairs");
            if (hpElem is null)
                return result;

            var vector = hpElem.Element(VtNs + "vector");
            if (vector is null)
                return result;

            var items = vector.Elements().ToList();
            for (int i = 0; i + 1 < items.Count; i += 2)
            {
                var nameElem = items[i];
                var countElem = items[i + 1];
                var name = nameElem.Value;
                int.TryParse(countElem.Value, CultureInfo.InvariantCulture, out var count);
                result.Add(new HeadingPairData(name, count));
            }

            return result;
        }
    }

    internal IReadOnlyList<string> TitlesOfParts
    {
        get
        {
            var result = new List<string>();
            var topElem = _root.Element(EpNs + "TitlesOfParts");
            if (topElem is null)
                return result;

            var vector = topElem.Element(VtNs + "vector");
            if (vector is null)
                return result;

            foreach (var item in vector.Elements())
                result.Add(item.Value);

            return result;
        }
    }

    internal XElement Root
    {
        get
        {
            SortIntoSchemaOrder();
            return _root;
        }
    }

    /// <summary>
    /// Rewrites the counts this part is supposed to describe rather than remember.
    /// </summary>
    /// <remarks>
    /// These are the numbers a file browser, a search indexer and PowerPoint's own info pane read
    /// to describe a deck without opening it. They are derived from the deck, so they are recomputed
    /// on every save; a part that ships stale or absent numbers is worse than one that ships none,
    /// because a reader has no way to tell that what it is being told is wrong.
    /// </remarks>
    internal void SetDerivedCounts(DeckStatistics statistics)
    {
        SetElement("Words", statistics.Words.ToString(CultureInfo.InvariantCulture));
        SetElement("Paragraphs", statistics.Paragraphs.ToString(CultureInfo.InvariantCulture));
        SetElement("Slides", statistics.Slides.ToString(CultureInfo.InvariantCulture));
        SetElement("Notes", statistics.Notes.ToString(CultureInfo.InvariantCulture));
        SetElement("HiddenSlides", statistics.HiddenSlides.ToString(CultureInfo.InvariantCulture));
        SetHeadingPairs(statistics);
        SetTitlesOfParts(statistics);
        MarkDirty();
    }

    /// <summary>
    /// Writes the "Theme"/"Slide Titles" heading pairs that index the TitlesOfParts vector.
    /// </summary>
    private void SetHeadingPairs(DeckStatistics statistics)
    {
        var vector = new XElement(VtNs + "vector",
            new XAttribute("size", 4),
            new XAttribute("baseType", "variant"),
            Variant(new XElement(VtNs + "lpstr", "Theme")),
            Variant(new XElement(VtNs + "i4", statistics.ThemeNames.Count)),
            Variant(new XElement(VtNs + "lpstr", "Slide Titles")),
            Variant(new XElement(VtNs + "i4", statistics.SlideTitles.Count)));

        ReplaceElement("HeadingPairs", vector);
    }

    /// <summary>
    /// Writes the theme names followed by the slide titles, in the order HeadingPairs indexes them.
    /// </summary>
    private void SetTitlesOfParts(DeckStatistics statistics)
    {
        var vector = new XElement(VtNs + "vector",
            new XAttribute("size", statistics.ThemeNames.Count + statistics.SlideTitles.Count),
            new XAttribute("baseType", "lpstr"));

        foreach (var name in statistics.ThemeNames)
            vector.Add(new XElement(VtNs + "lpstr", name));
        foreach (var title in statistics.SlideTitles)
            vector.Add(new XElement(VtNs + "lpstr", title));

        ReplaceElement("TitlesOfParts", vector);
    }

    private static XElement Variant(XElement content) => new(VtNs + "variant", content);

    private void ReplaceElement(string name, XElement content)
    {
        var element = _root.Element(EpNs + name);
        if (element is null)
        {
            element = new XElement(EpNs + name);
            _root.Add(element);
        }

        element.RemoveNodes();
        element.Add(content);
    }

    /// <summary>
    /// The order CT_Properties declares its children in (ECMA-376 Part 1, 15.2.12.2).
    /// </summary>
    /// <remarks>
    /// It is a sequence, not a choice: an element in the wrong place is a schema error even though
    /// every element present is a legal one. Properties are set in whatever order a caller happens
    /// to set them, so the order is imposed here, on the way out.
    /// </remarks>
    private static readonly string[] SchemaOrder =
    [
        "Template", "Manager", "Company", "Pages", "Words", "Characters", "PresentationFormat",
        "Lines", "Paragraphs", "Slides", "Notes", "TotalTime", "HiddenSlides", "MMClips",
        "ScaleCrop", "HeadingPairs", "TitlesOfParts", "LinksUpToDate", "CharactersWithSpaces",
        "SharedDoc", "HyperlinkBase", "HLinks", "HyperlinksChanged", "DigSig", "Application",
        "AppVersion", "DocSecurity",
    ];

    private void SortIntoSchemaOrder()
    {
        var children = _root.Elements().ToList();
        if (children.Count < 2)
            return;

        var ordered = children
            .OrderBy(e =>
            {
                var index = Array.IndexOf(SchemaOrder, e.Name.LocalName);
                return index < 0 ? SchemaOrder.Length : index;
            })
            .ToList();

        if (ordered.SequenceEqual(children))
            return;

        foreach (var child in children)
            child.Remove();
        foreach (var child in ordered)
            _root.Add(child);
    }

    /// <summary>
    /// Resets all writable app properties to defaults.
    /// </summary>
    internal void ClearAll()
    {
        Application = string.Empty;
        Company = string.Empty;
        Manager = string.Empty;
        PresentationFormat = string.Empty;
        Template = string.Empty;
        HyperlinkBase = string.Empty;
        TotalTime = TimeSpan.Zero;
        SharedDoc = false;
        ScaleCrop = false;
        LinksUpToDate = false;
        HyperlinksChanged = false;
    }

    private string? GetElement(string name) => _root.Element(EpNs + name)?.Value;

    private void SetElement(string name, string? value)
    {
        var elem = _root.Element(EpNs + name);
        if (elem is null)
        {
            elem = new XElement(EpNs + name);
            _root.Add(elem);
        }
        elem.Value = value ?? string.Empty;
    }

    private bool GetBoolElement(string name)
    {
        var text = GetElement(name);
        if (string.IsNullOrEmpty(text))
            return false;
        return text.Equals("true", StringComparison.OrdinalIgnoreCase) || text == "1";
    }

    private void SetBoolElement(string name, bool value)
    {
        SetElement(name, value ? "true" : "false");
    }

    private int GetIntElement(string name)
    {
        var text = GetElement(name);
        if (string.IsNullOrEmpty(text))
            return 0;
        return int.TryParse(text, CultureInfo.InvariantCulture, out var v) ? v : 0;
    }
}

/// <summary>
/// Data record for a heading pair entry.
/// </summary>
internal sealed record HeadingPairData(string Name, int Count);
