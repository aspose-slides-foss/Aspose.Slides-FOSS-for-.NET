using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of sections in a presentation.
/// </summary>
/// <remarks>
/// Sections are not part of the ECMA-376 presentation schema. PowerPoint stores them in an
/// extension on <c>ppt/presentation.xml</c> — a <c>&lt;p:ext&gt;</c> with the well-known URI
/// <c>{521415D9-36F7-43E2-AB2F-B90AF26B5E84}</c> holding a <c>&lt;p14:sectionLst&gt;</c>. A section
/// names the slides it contains by the numeric <c>id</c> of their <c>&lt;p:sldId&gt;</c> entries,
/// not by part name or relationship, so the section list only makes sense against the slide list it
/// was written with.
/// </remarks>
public sealed class SectionCollection : ISectionCollection
{
    private const string SectionListExtensionUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace P14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    private readonly List<ISection> _sections = [];

    private Presentation? _presentation;
    private PresentationPart? _presentationPart;
    private bool _parsed;
    private bool _dirty;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(Presentation presentation, PresentationPart presentationPart)
    {
        _presentation = presentation;
        _presentationPart = presentationPart;
    }

    /// <inheritdoc />
    public ISection this[int index] => Sections[index];

    /// <inheritdoc />
    public int Count => Sections.Count;

    /// <inheritdoc />
    public ISection AddSection(string name, ISlide startedFromSlide)
    {
        ArgumentNullException.ThrowIfNull(startedFromSlide);

        // A section is written as a range of the slide list of this presentation. A slide from
        // another one is not in that list, so the section would be written empty and the caller
        // would be told nothing.
        if (_presentation is not null && !SlideList().Contains(startedFromSlide))
        {
            throw new ArgumentException(
                "The slide a section starts from must be a slide of this presentation.",
                nameof(startedFromSlide));
        }

        var section = new Section(this)
        {
            Name = name,
            StartedFromSlide = startedFromSlide,
        };

        Sections.Add(section);
        MarkDirty();
        return section;
    }

    /// <inheritdoc />
    public ISection AppendEmptySection(string name, int startedFromIndex)
    {
        var slides = SlideList();
        var section = new Section(this)
        {
            Name = name,
            StartedFromSlide = startedFromIndex >= 0 && startedFromIndex < slides.Count
                ? slides[startedFromIndex]
                : null,
        };

        Sections.Add(section);
        MarkDirty();
        return section;
    }

    /// <inheritdoc />
    public int IndexOf(ISection section) => Sections.IndexOf(section);

    /// <inheritdoc />
    public void RemoveSection(ISection section)
    {
        if (Sections.Remove(section))
            MarkDirty();
    }

    /// <inheritdoc />
    public void RemoveSectionWithSlides(ISection section)
    {
        var slides = SlidesOf(section);

        if (!Sections.Remove(section))
            return;

        foreach (var slide in slides)
            _presentation?.Slides.Remove(slide);

        MarkDirty();
    }

    /// <inheritdoc />
    public void ReorderSectionWithSlides(ISection section, int index)
    {
        if (!Sections.Contains(section))
            return;

        // Capture the membership of every section before the order changes: membership is derived
        // from where each section starts in the slide list, so moving a section without moving its
        // slides would silently reassign them to whichever section now starts there.
        var membership = Sections.ToDictionary(s => s, SlidesOf);

        Sections.Remove(section);
        if (index >= Sections.Count)
            Sections.Add(section);
        else
            Sections.Insert(index, section);

        var order = new List<ISlide>();
        foreach (var s in Sections)
            order.AddRange(membership[s]);

        (_presentation?.Slides as SlideCollection)?.ReorderInternal(order);

        foreach (var s in Sections)
        {
            if (s is Section concrete && membership[s].Count > 0)
                concrete.StartedFromSlide = membership[s][0];
        }

        MarkDirty();
    }

    /// <inheritdoc />
    public void Clear()
    {
        if (Sections.Count == 0)
            return;

        Sections.Clear();
        MarkDirty();
    }

    /// <inheritdoc />
    public IEnumerator<ISection> GetEnumerator() => Sections.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    // ── Internal ─────────────────────────────────────────────────

    /// <summary>
    /// Records that the section list no longer matches what is in the package.
    /// </summary>
    internal void MarkDirty() => _dirty = true;

    /// <summary>
    /// Lists the slides belonging to <paramref name="section"/>: those from the slide it starts
    /// from up to the slide the next section starts from.
    /// </summary>
    internal List<ISlide> SlidesOf(ISection section)
    {
        var slides = SlideList();
        var boundaries = Boundaries(slides);

        var position = Sections.IndexOf(section);
        if (position < 0)
            return [];

        var start = boundaries[position];
        var end = position + 1 < boundaries.Count ? boundaries[position + 1] : slides.Count;

        // The first section owns anything in front of it: once a presentation has sections, a slide
        // outside every one of them has nowhere to appear in PowerPoint's section pane.
        if (position == 0)
            start = 0;

        return start >= end ? [] : slides.GetRange(start, end - start);
    }

    /// <summary>
    /// Writes the section list into <c>ppt/presentation.xml</c>, or removes it when there are no
    /// sections left.
    /// </summary>
    /// <remarks>
    /// Does nothing unless a mutator has run. A presentation opened and saved unchanged keeps the
    /// extension list byte for byte, including anything in it this library does not model.
    /// </remarks>
    internal void Flush()
    {
        if (!_dirty || _presentationPart is null)
            return;

        var root = _presentationPart.Document.Root;
        if (root is null)
            return;

        var extLst = root.Element(PNs + "extLst");
        var ext = extLst?.Elements(PNs + "ext")
            .FirstOrDefault(e => string.Equals(e.Attribute("uri")?.Value, SectionListExtensionUri,
                StringComparison.OrdinalIgnoreCase));

        if (_sections.Count == 0)
        {
            ext?.Remove();
            if (extLst is not null && !extLst.HasElements)
                extLst.Remove();
            _dirty = false;
            return;
        }

        if (extLst is null)
        {
            // CT_Presentation puts extLst last; everything else has already been written by now.
            extLst = new XElement(PNs + "extLst");
            root.Add(extLst);
        }

        if (ext is null)
        {
            ext = new XElement(PNs + "ext", new XAttribute("uri", SectionListExtensionUri));
            extLst.Add(ext);
        }

        ext.RemoveNodes();

        var slideIds = SlideIds();
        var sectionLst = new XElement(P14Ns + "sectionLst",
            new XAttribute(XNamespace.Xmlns + "p14", P14Ns.NamespaceName));

        foreach (var section in _sections)
        {
            var element = new XElement(P14Ns + "section",
                new XAttribute("name", section.Name),
                new XAttribute("id", FormatSectionId(section.SectionId)));

            var sldIdLst = new XElement(P14Ns + "sldIdLst");
            foreach (var slide in SlidesOf(section))
            {
                var index = SlideList().IndexOf(slide);
                if (index >= 0 && index < slideIds.Count)
                    sldIdLst.Add(new XElement(P14Ns + "sldId", new XAttribute("id", slideIds[index])));
            }

            element.Add(sldIdLst);
            sectionLst.Add(element);
        }

        ext.Add(sectionLst);
        _dirty = false;
    }

    // ── Private helpers ──────────────────────────────────────────

    private List<ISection> Sections
    {
        get
        {
            EnsureParsed();
            return _sections;
        }
    }

    private void EnsureParsed()
    {
        if (_parsed || _presentationPart is null || _presentation is null)
            return;

        _parsed = true;

        var sectionLst = _presentationPart.Document.Root
            ?.Element(PNs + "extLst")
            ?.Elements(PNs + "ext")
            .FirstOrDefault(e => string.Equals(e.Attribute("uri")?.Value, SectionListExtensionUri,
                StringComparison.OrdinalIgnoreCase))
            ?.Element(P14Ns + "sectionLst");

        if (sectionLst is null)
            return;

        var slides = SlideList();
        var slideIds = SlideIds();

        foreach (var element in sectionLst.Elements(P14Ns + "section"))
        {
            var firstId = element.Element(P14Ns + "sldIdLst")
                ?.Elements(P14Ns + "sldId")
                .FirstOrDefault()
                ?.Attribute("id")?.Value;

            var index = firstId is null ? -1 : slideIds.IndexOf(firstId);

            _sections.Add(new Section(this, ParseSectionId(element.Attribute("id")?.Value))
            {
                Name = element.Attribute("name")?.Value ?? string.Empty,
                StartedFromSlide = index >= 0 && index < slides.Count ? slides[index] : null,
            });
        }
    }

    /// <summary>
    /// The index in the slide list at which each section begins, in section order.
    /// </summary>
    private List<int> Boundaries(List<ISlide> slides)
    {
        var boundaries = new List<int>(_sections.Count);
        int previous = 0;

        foreach (var section in _sections)
        {
            var start = section.StartedFromSlide is null
                ? slides.Count
                : slides.IndexOf(section.StartedFromSlide);

            if (start < previous)
                start = previous;

            boundaries.Add(start);
            previous = start;
        }

        return boundaries;
    }

    private List<ISlide> SlideList()
    {
        return _presentation is null ? [] : [.. _presentation.Slides.AsIEnumerable];
    }

    /// <summary>
    /// The numeric <c>&lt;p:sldId&gt;</c> ids, in slide order — the identifiers a section names.
    /// </summary>
    private List<string> SlideIds()
    {
        var ids = new List<string>();
        var sldIdLst = _presentationPart?.Document.Root?.Element(PNs + "sldIdLst");
        if (sldIdLst is null)
            return ids;

        foreach (var sldId in sldIdLst.Elements(PNs + "sldId"))
        {
            var id = sldId.Attribute("id")?.Value;
            if (id is not null)
                ids.Add(id);
        }

        return ids;
    }

    private static string FormatSectionId(Guid id) => id.ToString("B").ToUpperInvariant();

    private static Guid ParseSectionId(string? value)
    {
        return Guid.TryParse(value, out var id) ? id : Guid.NewGuid();
    }
}
