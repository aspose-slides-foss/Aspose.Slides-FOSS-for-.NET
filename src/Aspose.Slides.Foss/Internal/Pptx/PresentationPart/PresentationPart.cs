using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.PresentationPart;

/// <summary>
/// Manages the ppt/presentation.xml part of an OPC package.
/// Provides methods to parse presentation structure, add/remove slide and master references,
/// and get/set presentation properties such as slide size and notes size.
/// </summary>
public sealed class PresentationPart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// The part name within the OPC package.
    /// </summary>
    public const string PartName = "ppt/presentation.xml";

    private readonly Internal.OpcPackage _package;
    private XElement _root = null!;
    private readonly List<SlideReference> _slideRefs = [];
    private readonly List<MasterReference> _masterRefs = [];

    /// <summary>
    /// Initializes the presentation part manager and loads presentation.xml from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation.</param>
    internal PresentationPart(Internal.OpcPackage package)
    {
        _package = package;
        Load();
    }

    /// <summary>
    /// Loads and parses the presentation.xml from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(PartName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            var doc = XDocument.Load(ms);
            _root = doc.Root ?? throw new InvalidOperationException("presentation.xml has no root element");
            ParseSlides();
            ParseMasters();
        }
        else
        {
            throw new InvalidOperationException($"{PartName} not found in package");
        }
    }

    /// <summary>
    /// Parses slide references from the presentation XML.
    /// </summary>
    public void ParseSlides()
    {
        _slideRefs.Clear();

        var sldIdLst = _root.Descendants(PNs + "sldIdLst").FirstOrDefault();
        if (sldIdLst is null) return;

        foreach (var sldIdElem in sldIdLst.Elements(PNs + "sldId"))
        {
            var slideId = (int?)sldIdElem.Attribute("id") ?? 0;
            var relId = (string?)sldIdElem.Attribute(RNs + "id") ?? "";
            _slideRefs.Add(new SlideReference(slideId, relId));
        }
    }

    /// <summary>
    /// Parses master slide references from the presentation XML.
    /// </summary>
    public void ParseMasters()
    {
        _masterRefs.Clear();

        var sldMasterIdLst = _root.Descendants(PNs + "sldMasterIdLst").FirstOrDefault();
        if (sldMasterIdLst is null) return;

        foreach (var sldMasterIdElem in sldMasterIdLst.Elements(PNs + "sldMasterId"))
        {
            var masterId = (int?)sldMasterIdElem.Attribute("id") ?? 0;
            var relId = (string?)sldMasterIdElem.Attribute(RNs + "id") ?? "";
            _masterRefs.Add(new MasterReference(masterId, relId));
        }
    }

    /// <summary>
    /// Gets a copy of the master slide references list.
    /// </summary>
    public List<MasterReference> MasterReferences => [.. _masterRefs];

    /// <summary>
    /// Gets the root element of the presentation XML.
    /// </summary>
    public XElement? ElementTree => _root;

    /// <summary>
    /// Gets a copy of the slide references list in order.
    /// </summary>
    public List<SlideReference> SlideReferences => [.. _slideRefs];

    /// <summary>
    /// Gets the number of slides.
    /// </summary>
    public int SlideCount => _slideRefs.Count;

    /// <summary>
    /// Gets a slide reference by its unique slide ID.
    /// </summary>
    /// <param name="slideId">The unique slide ID.</param>
    /// <returns>The matching <see cref="SlideReference"/>, or <c>null</c> if not found.</returns>
    public SlideReference? GetSlideRefById(int slideId)
    {
        return _slideRefs.FirstOrDefault(r => r.SlideId == slideId);
    }

    /// <summary>
    /// Gets a slide reference by its relationship ID.
    /// </summary>
    /// <param name="relId">The relationship ID (e.g., "rId2").</param>
    /// <returns>The matching <see cref="SlideReference"/>, or <c>null</c> if not found.</returns>
    public SlideReference? GetSlideRefByRelId(string relId)
    {
        return _slideRefs.FirstOrDefault(r => r.RelId == relId);
    }

    /// <summary>
    /// Generates the next available slide ID. Slide IDs start at 256 by PPTX convention.
    /// </summary>
    public int GetNextSlideId()
    {
        if (_slideRefs.Count == 0)
            return 256;
        return _slideRefs.Max(r => r.SlideId) + 1;
    }

    /// <summary>
    /// Generates the next available master/layout slide ID.
    /// Master IDs and layout IDs share the same ID space and must all be unique.
    /// </summary>
    public int GetNextMasterId()
    {
        var maxId = 2147483647;

        foreach (var r in _masterRefs)
        {
            if (r.MasterId > maxId)
                maxId = r.MasterId;
        }

        // Check existing layout IDs across all master slides in the package
        foreach (var partName in _package.GetPartNames())
        {
            if (partName.StartsWith("ppt/slideMasters/", StringComparison.OrdinalIgnoreCase)
                && partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                var content = _package.GetPart(partName);
                if (content is not null)
                {
                    using var msContent = new MemoryStream(content);
                    var doc = XDocument.Load(msContent);
                    foreach (var elem in doc.Descendants(PNs + "sldLayoutId"))
                    {
                        var layoutId = (int?)elem.Attribute("id") ?? 0;
                        if (layoutId > maxId)
                            maxId = layoutId;
                    }
                }
            }
        }

        return maxId + 1;
    }

    /// <summary>
    /// Adds a new master slide reference to the presentation.
    /// </summary>
    /// <param name="relId">The relationship ID for the master slide.</param>
    /// <param name="masterId">Optional specific master ID. If <c>null</c>, auto-generated.</param>
    /// <returns>The created <see cref="MasterReference"/>.</returns>
    public MasterReference AddMasterReference(string relId, int? masterId = null)
    {
        masterId ??= GetNextMasterId();

        var r = new MasterReference(masterId.Value, relId);
        _masterRefs.Add(r);

        var sldMasterIdLst = _root.Descendants(PNs + "sldMasterIdLst").FirstOrDefault();
        if (sldMasterIdLst is null)
        {
            var sldIdLst = _root.Descendants(PNs + "sldIdLst").FirstOrDefault();
            sldMasterIdLst = new XElement(PNs + "sldMasterIdLst");
            if (sldIdLst is not null)
            {
                sldIdLst.AddBeforeSelf(sldMasterIdLst);
            }
            else
            {
                _root.Add(sldMasterIdLst);
            }
        }

        var sldMasterIdElem = new XElement(PNs + "sldMasterId",
            new XAttribute("id", masterId.Value),
            new XAttribute(RNs + "id", relId));
        sldMasterIdLst.Add(sldMasterIdElem);

        return r;
    }

    /// <summary>
    /// Adds a new slide reference to the presentation.
    /// </summary>
    /// <param name="relId">The relationship ID for the slide.</param>
    /// <param name="slideId">Optional specific slide ID. If <c>null</c>, auto-generated.</param>
    /// <param name="index">Position to insert at. -1 means append at end.</param>
    /// <returns>The created <see cref="SlideReference"/>.</returns>
    public SlideReference AddSlideReference(string relId, int? slideId = null, int index = -1)
    {
        slideId ??= GetNextSlideId();

        var r = new SlideReference(slideId.Value, relId);

        if (index < 0 || index >= _slideRefs.Count)
            _slideRefs.Add(r);
        else
            _slideRefs.Insert(index, r);

        var sldIdLst = _root.Descendants(PNs + "sldIdLst").FirstOrDefault();
        if (sldIdLst is null)
        {
            sldIdLst = new XElement(PNs + "sldIdLst");
            _root.Add(sldIdLst);
        }

        var sldIdElem = new XElement(PNs + "sldId",
            new XAttribute("id", slideId.Value),
            new XAttribute(RNs + "id", relId));

        var children = sldIdLst.Elements().ToList();
        if (index < 0 || index >= children.Count)
        {
            sldIdLst.Add(sldIdElem);
        }
        else
        {
            children[index].AddBeforeSelf(sldIdElem);
        }

        return r;
    }

    /// <summary>
    /// Removes a slide reference by its slide ID.
    /// </summary>
    /// <param name="slideId">The unique slide ID to remove.</param>
    /// <returns><c>true</c> if removed, <c>false</c> if not found.</returns>
    public bool RemoveSlideReference(int slideId)
    {
        var idx = _slideRefs.FindIndex(r => r.SlideId == slideId);
        if (idx < 0)
            return false;

        _slideRefs.RemoveAt(idx);

        var sldIdLst = _root.Descendants(PNs + "sldIdLst").FirstOrDefault();
        sldIdLst?.Elements(PNs + "sldId")
            .FirstOrDefault(e => (string?)e.Attribute("id") == slideId.ToString())
            ?.Remove();

        return true;
    }

    /// <summary>
    /// Gets the slide size in EMUs (English Metric Units).
    /// </summary>
    /// <returns>A tuple of (width, height) in EMUs.</returns>
    public (int Width, int Height) GetSlideSize()
    {
        var sldSz = _root.Descendants(PNs + "sldSz").FirstOrDefault();
        if (sldSz is not null)
        {
            var cx = (int?)sldSz.Attribute("cx") ?? 9144000;
            var cy = (int?)sldSz.Attribute("cy") ?? 6858000;
            return (cx, cy);
        }
        return (9144000, 6858000);
    }

    /// <summary>
    /// Gets the notes slide size in EMUs.
    /// </summary>
    /// <returns>A tuple of (width, height) in EMUs.</returns>
    public (int Width, int Height) GetNotesSize()
    {
        var notesSz = _root.Descendants(PNs + "notesSz").FirstOrDefault();
        if (notesSz is not null)
        {
            var cx = (int?)notesSz.Attribute("cx") ?? 6858000;
            var cy = (int?)notesSz.Attribute("cy") ?? 9144000;
            return (cx, cy);
        }
        return (6858000, 9144000);
    }

    /// <summary>
    /// Sets the notes slide size in EMUs.
    /// </summary>
    /// <param name="cx">Width in EMUs.</param>
    /// <param name="cy">Height in EMUs.</param>
    public void SetNotesSize(int cx, int cy)
    {
        var notesSz = _root.Descendants(PNs + "notesSz").FirstOrDefault();
        if (notesSz is not null)
        {
            notesSz.SetAttributeValue("cx", cx);
            notesSz.SetAttributeValue("cy", cy);
        }
        else
        {
            notesSz = new XElement(PNs + "notesSz",
                new XAttribute("cx", cx),
                new XAttribute("cy", cy));

            var sldSz = _root.Descendants(PNs + "sldSz").FirstOrDefault();
            if (sldSz is not null)
            {
                sldSz.AddAfterSelf(notesSz);
            }
            else
            {
                _root.Add(notesSz);
            }
        }
    }

    /// <summary>
    /// Gets the first slide number for numbering.
    /// </summary>
    /// <returns>The first slide number, defaulting to 1.</returns>
    public int GetFirstSlideNumber()
    {
        return (int?)_root.Attribute("firstSlideNum") ?? 1;
    }

    /// <summary>
    /// Sets the first slide number for numbering.
    /// </summary>
    /// <param name="number">The first slide number.</param>
    public void SetFirstSlideNumber(int number)
    {
        _root.SetAttributeValue("firstSlideNum", number);
    }

    /// <summary>
    /// Saves the presentation.xml back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());
    }
}
