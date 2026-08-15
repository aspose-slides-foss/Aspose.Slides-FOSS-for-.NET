using System.Collections;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of slides in a presentation.
/// </summary>
public sealed class SlideCollection : ISlideCollection
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string SlideRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    private const string SlideLayoutRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    private Presentation? _presentation;
    private OpcPackage? _package;
    private PresentationPart? _presentationPart;
    private Func<string, ILayoutSlide?>? _layoutResolver;
    private bool _parsed;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        Presentation presentation,
        OpcPackage package,
        PresentationPart presentationPart,
        Func<string, ILayoutSlide?> layoutResolver)
    {
        _presentation = presentation;
        _package = package;
        _presentationPart = presentationPart;
        _layoutResolver = layoutResolver;
    }

    private List<ISlide> Slides
    {
        get
        {
            EnsureParsed();
            return _presentation!.SlidesInternal;
        }
    }

    /// <inheritdoc />
    public ISlide this[int index] => Slides[index];

    /// <inheritdoc />
    public int Count => Slides.Count;

    /// <inheritdoc />
    public IList<ISlide> AsICollection => Slides;

    /// <inheritdoc />
    public IEnumerable<ISlide> AsIEnumerable => Slides;

    /// <inheritdoc />
    public ISlide AddClone(ISlide sourceSlide)
    {
        var slide = CloneSlide(sourceSlide);
        Slides.Add(slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide AddClone(ISlide sourceSlide, ILayoutSlide destLayout)
    {
        var slide = CloneSlide(sourceSlide);
        slide.LayoutSlide = destLayout;
        Slides.Add(slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide AddClone(ISlide sourceSlide, IMasterSlide destMaster, bool allowCloneMissingLayout)
    {
        var slide = CloneSlide(sourceSlide);
        Slides.Add(slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide InsertClone(int index, ISlide sourceSlide)
    {
        var slide = CloneSlide(sourceSlide);
        Slides.Insert(index, slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide InsertClone(int index, ISlide sourceSlide, ILayoutSlide destLayout)
    {
        var slide = CloneSlide(sourceSlide);
        slide.LayoutSlide = destLayout;
        Slides.Insert(index, slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide InsertClone(int index, ISlide sourceSlide, IMasterSlide destMaster, bool allowCloneMissingLayout)
    {
        var slide = CloneSlide(sourceSlide);
        Slides.Insert(index, slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide[] ToArray() => [.. Slides];

    /// <inheritdoc />
    public ISlide[] ToArray(int startIndex, int count)
    {
        return Slides.GetRange(startIndex, count).ToArray();
    }

    /// <inheritdoc />
    public ISlide AddEmptySlide(ILayoutSlide layout)
    {
        var slide = CreateNewSlide(layout);
        Slides.Add(slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public ISlide InsertEmptySlide(int index, ILayoutSlide layout)
    {
        var slide = CreateNewSlide(layout);
        Slides.Insert(index, slide);
        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <inheritdoc />
    public void Remove(ISlide value)
    {
        Slides.Remove(value);
        UpdateSlideNumbers();
        PersistSlideList();
    }

    /// <inheritdoc />
    public void RemoveAt(int index)
    {
        Slides.RemoveAt(index);
        UpdateSlideNumbers();
        PersistSlideList();
    }

    /// <inheritdoc />
    public int IndexOf(ISlide slide)
    {
        var slides = Slides;
        for (int i = 0; i < slides.Count; i++)
        {
            if (ReferenceEquals(slides[i], slide))
                return i;
        }

        return -1;
    }

    /// <inheritdoc />
    public IEnumerator<ISlide> GetEnumerator() => Slides.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    // ── Internal methods ─────────────────────────────────────────

    /// <summary>
    /// Replaces the slide order with <paramref name="order"/> and rewrites
    /// <c>&lt;p:sldIdLst&gt;</c> to match.
    /// </summary>
    /// <param name="order">
    /// The slides in their new order. Slides of this presentation the list does not mention follow
    /// it, keeping their relative order, so a partial order can never drop a slide.
    /// </param>
    internal void ReorderInternal(IEnumerable<ISlide> order)
    {
        var current = Slides;
        var reordered = new List<ISlide>();

        foreach (var slide in order)
        {
            if (current.Contains(slide) && !reordered.Contains(slide))
                reordered.Add(slide);
        }

        foreach (var slide in current)
        {
            if (!reordered.Contains(slide))
                reordered.Add(slide);
        }

        current.Clear();
        current.AddRange(reordered);

        UpdateSlideNumbers();
        PersistSlideList();
    }

    /// <summary>
    /// Finds the next available slide file number by scanning existing part names.
    /// </summary>
    internal int GetNextSlideFileNumber()
    {
        var existingNums = new HashSet<int>();
        foreach (var partName in _package!.GetPartNames())
        {
            var m = Regex.Match(partName, @"^ppt/slides/slide(\d+)\.xml$");
            if (m.Success)
                existingNums.Add(int.Parse(m.Groups[1].Value));
        }

        int num = 1;
        while (existingNums.Contains(num))
            num++;
        return num;
    }

    /// <summary>
    /// Internal implementation for adding or inserting an empty slide.
    /// </summary>
    /// <param name="layout">The layout slide to use.</param>
    /// <param name="index">Position to insert at. -1 means append at end.</param>
    /// <returns>The newly created slide.</returns>
    internal ISlide AddEmptySlideInternal(ILayoutSlide layout, int index = -1)
    {
        var slide = CreateNewSlide(layout);

        if (index < 0 || index >= Slides.Count)
            Slides.Add(slide);
        else
            Slides.Insert(index, slide);

        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <summary>
    /// Internal implementation for cloning a slide with optional layout and master resolution.
    /// </summary>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <param name="index">Position to insert at. -1 means append at end.</param>
    /// <param name="destLayout">Optional destination layout slide.</param>
    /// <param name="destMaster">Optional destination master slide.</param>
    /// <param name="allowCloneMissingLayout">If true and layout not found, use a fallback layout.</param>
    /// <returns>The cloned slide.</returns>
    internal ISlide CloneSlideInternal(
        ISlide sourceSlide,
        int index = -1,
        ILayoutSlide? destLayout = null,
        IMasterSlide? destMaster = null,
        bool allowCloneMissingLayout = false)
    {
        var slide = CloneSlide(sourceSlide);

        // Determine the layout to use
        if (destLayout is not null)
        {
            slide.LayoutSlide = destLayout;
        }
        else if (destMaster is not null)
        {
            var layoutPartName = FindMatchingLayout(sourceSlide, destMaster, allowCloneMissingLayout);
            if (layoutPartName is not null)
                slide.LayoutSlide = _layoutResolver?.Invoke(layoutPartName);
        }
        else
        {
            // Check if cloning within the same presentation
            var isSamePackage = sourceSlide is Slide src && ReferenceEquals(src.PackageInternal, _package);
            if (isSamePackage)
            {
                slide.LayoutSlide = sourceSlide.LayoutSlide;
            }
            else if (sourceSlide.LayoutSlide is ILayoutSlide srcLayout)
            {
                var layoutPartName = CloneMasterChainForSlide(sourceSlide, srcLayout);
                if (layoutPartName is not null)
                    slide.LayoutSlide = _layoutResolver?.Invoke(layoutPartName);
            }
        }

        if (index < 0 || index >= Slides.Count)
            Slides.Add(slide);
        else
            Slides.Insert(index, slide);

        UpdateSlideNumbers();
        PersistSlideList();
        return slide;
    }

    /// <summary>
    /// Finds a matching layout in the destination master for the source slide.
    /// </summary>
    internal string? FindMatchingLayout(ISlide sourceSlide, IMasterSlide destMaster, bool allowClone)
    {
        var sourceLayout = sourceSlide.LayoutSlide;
        if (sourceLayout is null)
            return null;

        var sourceLayoutType = sourceLayout.LayoutType;

        // Match by layout type
        foreach (var layout in destMaster.LayoutSlides)
        {
            if (layout.LayoutType == sourceLayoutType)
                return GetLayoutPartName(layout);
        }

        // Match by name
        var sourceLayoutName = sourceLayout.Name;
        foreach (var layout in destMaster.LayoutSlides)
        {
            if (layout.Name == sourceLayoutName)
                return GetLayoutPartName(layout);
        }

        // Fallback to first layout
        foreach (var layout in destMaster.LayoutSlides)
            return GetLayoutPartName(layout);

        return null;
    }

    /// <summary>
    /// Finds a layout in the destination presentation matching the source layout type.
    /// </summary>
    internal string? FindLayoutByType(ILayoutSlide sourceLayout)
    {
        SlideLayoutType sourceType;
        try
        {
            sourceType = sourceLayout.LayoutType;
        }
        catch
        {
            return FindLayoutFromLayoutSlides(sourceLayout);
        }

        // Search presentation layout slides
        try
        {
            foreach (var layout in _presentation!.LayoutSlides)
            {
                try
                {
                    if (layout.LayoutType == sourceType)
                        return GetLayoutPartName(layout);
                }
                catch
                {
                    // Skip layouts with inaccessible type
                }
            }
        }
        catch
        {
            // LayoutSlides not available
        }

        return GetFirstLayoutPartName();
    }

    /// <summary>
    /// Finds a matching layout by name using layout slides.
    /// </summary>
    internal string? FindLayoutFromLayoutSlides(ILayoutSlide sourceLayout)
    {
        try
        {
            var sourceName = sourceLayout.Name;
            foreach (var layout in _presentation!.LayoutSlides)
            {
                try
                {
                    if (layout.Name == sourceName)
                        return GetLayoutPartName(layout);
                }
                catch
                {
                    // Skip layouts with inaccessible name
                }
            }
        }
        catch
        {
            // LayoutSlides not available
        }

        return GetFirstLayoutPartName();
    }

    /// <summary>
    /// Gets the first available layout part name in the destination presentation.
    /// </summary>
    internal string? GetFirstLayoutPartName()
    {
        // Try layout slides from presentation
        try
        {
            foreach (var layout in _presentation!.LayoutSlides)
                return GetLayoutPartName(layout);
        }
        catch
        {
            // LayoutSlides not available
        }

        // Last resort: scan package for layout files
        foreach (var partName in _package!.GetPartNames())
        {
            if (partName.StartsWith("ppt/slideLayouts/", StringComparison.Ordinal) && partName.EndsWith(".xml", StringComparison.Ordinal))
                return partName;
        }

        return null;
    }

    /// <summary>
    /// Clones the master slide chain for a slide being cloned from another presentation.
    /// </summary>
    /// <param name="sourceSlide">The source slide being cloned.</param>
    /// <param name="sourceLayout">The source slide's layout.</param>
    /// <returns>The part name of the cloned layout to use for the new slide.</returns>
    internal string? CloneMasterChainForSlide(ISlide sourceSlide, ILayoutSlide sourceLayout)
    {
        // Get the source master
        IMasterSlide? sourceMaster;
        try
        {
            sourceMaster = sourceLayout.MasterSlide;
        }
        catch
        {
            return GetFirstLayoutPartName();
        }

        if (sourceMaster is null)
            return GetFirstLayoutPartName();

        // Clone the master to the destination presentation
        IMasterSlide clonedMaster;
        try
        {
            clonedMaster = _presentation!.Masters.AddClone(sourceMaster);
        }
        catch
        {
            return GetFirstLayoutPartName();
        }

        // Find the cloned layout matching the source layout
        SlideLayoutType? sourceLayoutType = null;
        string? sourceLayoutName = null;
        try { sourceLayoutType = sourceLayout.LayoutType; } catch { }
        try { sourceLayoutName = sourceLayout.Name; } catch { }

        foreach (var clonedLayout in clonedMaster.LayoutSlides)
        {
            // Match by type first
            if (sourceLayoutType is not null)
            {
                try
                {
                    if (clonedLayout.LayoutType == sourceLayoutType)
                        return GetLayoutPartName(clonedLayout);
                }
                catch { }
            }

            // Match by name
            if (sourceLayoutName is not null)
            {
                try
                {
                    if (clonedLayout.Name == sourceLayoutName)
                        return GetLayoutPartName(clonedLayout);
                }
                catch { }
            }
        }

        // Fallback to first layout from cloned master
        foreach (var layout in clonedMaster.LayoutSlides)
            return GetLayoutPartName(layout);

        return GetFirstLayoutPartName();
    }

    private static string? GetLayoutPartName(ILayoutSlide layout)
    {
        return layout is LayoutSlide ls ? ls.InternalPartName : null;
    }

    // ── Private helpers ──────────────────────────────────────────

    private void EnsureParsed()
    {
        if (_parsed || _presentation is null)
            return;

        _parsed = true;

        var entries = _presentationPart!.GetSlideEntries();
        foreach (var (slideId, relId) in entries)
        {
            var rel = _presentationPart.Rels.GetById(relId);
            if (rel is null)
                continue;

            var partName = ResolvePartName("ppt/", rel.Target);
            var slide = new Slide();
            slide.PartName = partName;
            slide.PackageInternal = _package;
            slide.SetPresentation(_presentation);
            slide.SlideNumber = _presentation.SlidesInternal.Count + 1;

            // Load slide XML and init slide part for shapes
            var slideData = _package!.GetPart(partName);
            SlidePart? slidePart = null;
            if (slideData is not null)
            {
                var slideDoc = XDocument.Parse(System.Text.Encoding.UTF8.GetString(slideData));
                slidePart = new SlidePart();
                slidePart.InitInternal(partName);
                slidePart.Element = slideDoc.Root;
                slidePart.Package = _package;
                slide.SetSlidePart(slidePart);
            }

            // Resolve layout from slide rels and load rels into SlidePart
            var slideRelsName = GetRelsPath(partName);
            var slideRelsData = _package.GetPart(slideRelsName);
            if (slideRelsData is not null)
            {
                // Load rels into the SlidePart so NotesSlideManager can find notes relationships
                slidePart?.RelsManager.Load(slideRelsData);

                var slideRels = slidePart?.RelsManager ?? new RelsManager();
                if (slidePart is null)
                    slideRels.Load(slideRelsData);

                foreach (var layoutRel in slideRels.FindByType(SlideLayoutRelType))
                {
                    var layoutPartName = ResolvePartName(
                        GetDirectoryPath(partName), layoutRel.Target);
                    slide.LayoutSlide = _layoutResolver?.Invoke(layoutPartName);
                    break;
                }
            }

            _presentation.SlidesInternal.Add(slide);
        }
    }

    private Slide CreateNewSlide(ILayoutSlide? layout)
    {
        var partName = AllocateSlidePartName();
        var slidePart = StoreSlidePart(partName, BuildEmptySlideXml());

        var slide = new Slide();
        slide.PartName = partName;
        slide.PackageInternal = _package;
        slide.SetPresentation(_presentation);
        slide.LayoutSlide = layout;
        slide.SetSlidePart(slidePart);

        return slide;
    }

    private Slide CloneSlide(ISlide sourceSlide)
    {
        var partName = AllocateSlidePartName();

        // Clone slide XML from source
        XDocument slideXml;
        if (sourceSlide is Slide src && src.GetSlidePartInternal()?.Element is not null)
        {
            slideXml = new XDocument(new XElement(src.GetSlidePartInternal()!.Element!));
        }
        else
        {
            slideXml = BuildEmptySlideXml();
        }

        var slidePart = StoreSlidePart(partName, slideXml);

        var slide = new Slide();
        slide.PartName = partName;
        slide.PackageInternal = _package;
        slide.SetPresentation(_presentation);
        slide.LayoutSlide = sourceSlide.LayoutSlide;
        slide.SetSlidePart(slidePart);

        return slide;
    }

    /// <summary>
    /// Picks a slide part name that is not already taken in the package.
    /// </summary>
    private string AllocateSlidePartName()
    {
        var slideIndex = Slides.Count + 1;
        var partName = $"ppt/slides/slide{slideIndex}.xml";

        while (_package!.GetPart(partName) is not null)
        {
            slideIndex++;
            partName = $"ppt/slides/slide{slideIndex}.xml";
        }

        return partName;
    }

    /// <summary>
    /// Writes a new slide part into the package with everything that makes it a slide: the markup,
    /// its content-type <c>Override</c> and its relationship to a layout.
    /// </summary>
    /// <remarks>
    /// The <c>Override</c> is not optional bookkeeping. Without it the part resolves through
    /// <c>&lt;Default Extension="xml"/&gt;</c> to <c>application/xml</c>, and a reader that checks
    /// content types — the Open XML SDK, python-pptx — refuses the package rather than the part.
    /// </remarks>
    private SlidePart StoreSlidePart(string partName, XDocument slideXml)
    {
        using var ms = new MemoryStream();
        slideXml.Save(ms);
        _package!.SetPart(partName, ms.ToArray());

        OpcRegistration.AddContentTypeOverride(_package, partName, PartContentTypes.Slide);

        var slidePart = new SlidePart();
        slidePart.InitInternal(partName);
        slidePart.Element = slideXml.Root;
        slidePart.Package = _package;
        slidePart.RelsManager.Add(SlideLayoutRelType, "../slideLayouts/slideLayout1.xml");
        slidePart.RelsManager.Save();

        return slidePart;
    }

    private void UpdateSlideNumbers()
    {
        for (int i = 0; i < Slides.Count; i++)
        {
            if (Slides[i] is Slide slide)
                slide.SlideNumber = i + 1;
        }
    }

    private void PersistSlideList()
    {
        if (_presentationPart is null)
            return;

        var root = _presentationPart.Document.Root;
        if (root is null)
            return;

        // Remove existing sldIdLst
        root.Element(PNs + "sldIdLst")?.Remove();

        // Build new sldIdLst
        var sldIdLst = new XElement(PNs + "sldIdLst");
        uint baseId = 256;
        foreach (var slide in Slides)
        {
            if (slide is not Slide s)
                continue;

            var relTarget = MakeRelativeTarget(s.PartName);
            var relId = FindOrAddSlideRel(relTarget);

            sldIdLst.Add(new XElement(PNs + "sldId",
                new XAttribute("id", baseId++),
                new XAttribute(RNs + "id", relId)));
        }

        // Insert sldIdLst after sldMasterIdLst (or at beginning)
        var masterIdLst = root.Element(PNs + "sldMasterIdLst");
        if (masterIdLst is not null)
            masterIdLst.AddAfterSelf(sldIdLst);
        else
            root.AddFirst(sldIdLst);

        _presentationPart.Flush();
    }

    private string FindOrAddSlideRel(string target)
    {
        foreach (var rel in _presentationPart!.Rels.FindByType(SlideRelType))
        {
            if (string.Equals(rel.Target, target, StringComparison.OrdinalIgnoreCase))
                return rel.Id;
        }

        return _presentationPart.Rels.Add(SlideRelType, target);
    }

    private static string MakeRelativeTarget(string partName)
    {
        // Convert "ppt/slides/slide1.xml" to "slides/slide1.xml"
        const string prefix = "ppt/";
        return partName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            ? partName[prefix.Length..]
            : partName;
    }

    private static XDocument BuildEmptySlideXml()
    {
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(PNs + "sld",
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
                        new XElement(PNs + "grpSpPr")))));
    }

    private static string ResolvePartName(string basePath, string target)
    {
        var combined = basePath + target;
        var segments = combined.Split('/');
        var resolved = new Stack<string>();

        foreach (var segment in segments)
        {
            if (segment == "..")
            {
                if (resolved.Count > 0)
                    resolved.Pop();
            }
            else if (segment != "." && segment.Length > 0)
            {
                resolved.Push(segment);
            }
        }

        return string.Join("/", resolved.Reverse());
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
