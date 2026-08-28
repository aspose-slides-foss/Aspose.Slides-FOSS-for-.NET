using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// What <c>docProps/app.xml</c> says about a deck, measured from the deck.
/// </summary>
internal sealed record DeckStatistics(
    int Slides,
    int HiddenSlides,
    int Notes,
    int Paragraphs,
    int Words,
    IReadOnlyList<string> ThemeNames,
    IReadOnlyList<string> SlideTitles)
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly char[] WordSeparators = [' ', '\t', '\r', '\n', '\f', '\v'];

    /// <summary>
    /// Measures a package that has already had everything held in memory written back into it.
    /// </summary>
    /// <param name="package">The package about to be serialized.</param>
    /// <param name="hiddenSlides">
    /// How many slides are hidden. This is read from the object model rather than the markup
    /// because <c>show="0"</c> is the absence of a default, and counting it from XML would need the
    /// same defaulting rules the model already applies.
    /// </param>
    internal static DeckStatistics Collect(OpcPackage package, int hiddenSlides)
    {
        var slideParts = PackageSlides.ListSlideParts(package);

        int paragraphs = 0;
        int words = 0;
        var titles = new List<string>();

        foreach (var slidePartName in slideParts)
        {
            var slide = Parse(package, slidePartName);
            if (slide is null)
            {
                titles.Add(string.Empty);
                continue;
            }

            foreach (var paragraph in slide.Descendants(ANs + "p"))
            {
                var text = TextOf(paragraph);
                if (text.Length == 0)
                    continue;

                paragraphs++;
                words += text.Split(WordSeparators, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            titles.Add(TitleOf(slide));
        }

        var notes = package.GetPartNames()
            .Count(name => name.StartsWith("ppt/notesSlides/", StringComparison.OrdinalIgnoreCase)
                        && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

        return new DeckStatistics(
            Slides: slideParts.Count,
            HiddenSlides: hiddenSlides,
            Notes: notes,
            Paragraphs: paragraphs,
            Words: words,
            ThemeNames: ThemeNamesIn(package),
            SlideTitles: titles);
    }

    private static XElement? Parse(OpcPackage package, string partName)
    {
        var data = package.GetPart(partName);
        if (data is null)
            return null;

        using var ms = new MemoryStream(data);
        return XDocument.Load(ms).Root;
    }

    private static string TextOf(XElement paragraph)
    {
        return string.Concat(paragraph.Descendants(ANs + "t").Select(t => t.Value)).Trim();
    }

    /// <summary>
    /// The text of the slide's title placeholder, or an empty entry when it has none.
    /// </summary>
    private static string TitleOf(XElement slide)
    {
        foreach (var shape in slide.Descendants(PNs + "sp"))
        {
            var placeholderType = shape.Element(PNs + "nvSpPr")
                ?.Element(PNs + "nvPr")
                ?.Element(PNs + "ph")
                ?.Attribute("type")?.Value;

            if (placeholderType is "title" or "ctrTitle")
            {
                var body = shape.Element(PNs + "txBody");
                if (body is not null)
                    return string.Concat(body.Descendants(ANs + "t").Select(t => t.Value)).Trim();
            }
        }

        return string.Empty;
    }

    private static List<string> ThemeNamesIn(OpcPackage package)
    {
        var names = new List<string>();

        foreach (var partName in package.GetSortedPartNames())
        {
            if (!partName.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase))
                continue;

            var theme = Parse(package, partName);
            names.Add(theme?.Attribute("name")?.Value ?? string.Empty);
        }

        return names;
    }
}
