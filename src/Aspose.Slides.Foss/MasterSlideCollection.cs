using System.Collections;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of master slides in a presentation.
/// </summary>
public sealed class MasterSlideCollection : IMasterSlideCollection, IEnumerable<IMasterSlide>
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string SlideMasterRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    private const string SlideLayoutRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    private Presentation? _presentation;
    private OpcPackage? _package;
    private PresentationPart? _presentationPart;
    private List<IMasterSlide>? _masters;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        Presentation presentation,
        OpcPackage package,
        PresentationPart presentationPart,
        List<IMasterSlide> masters)
    {
        _presentation = presentation;
        _package = package;
        _presentationPart = presentationPart;
        _masters = masters;
    }

    private List<IMasterSlide> Masters => _masters ??= [];

    /// <inheritdoc />
    public IMasterSlide this[int index] => Masters[index];

    /// <inheritdoc />
    public int Count => Masters.Count;

    /// <inheritdoc />
    public IList<IMasterSlide> AsICollection => Masters;

    /// <inheritdoc />
    public IEnumerable<IMasterSlide> AsIEnumerable => Masters;

    /// <summary>
    /// Returns an enumerator that iterates through the master slides.
    /// </summary>
    public IEnumerator<IMasterSlide> GetEnumerator() => Masters.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public IMasterSlide AddClone(IMasterSlide sourceMaster)
    {
        if (sourceMaster is not MasterSlide source)
        {
            var empty = new MasterSlide();
            Masters.Add(empty);
            return empty;
        }

        var sourcePartName = source.PartName;
        var nextMasterNum = GetNextMasterFileNumber();
        var destMasterPartName = $"ppt/slideMasters/slideMaster{nextMasterNum}.xml";

        // Clone master slide XML
        ClonePartXml(sourcePartName, destMasterPartName);

        // Read source master's rels to find layouts and other resources
        var sourceRels = LoadRels(sourcePartName);
        var destMasterRels = new RelsManager();
        var layoutPartMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Clone non-layout relationships (theme, images, etc.) and track layout rels
        foreach (var rel in sourceRels.FindByType(SlideLayoutRelType))
        {
            var sourceLayoutPartName = ResolvePartName(
                GetDirectoryPath(sourcePartName), rel.Target);
            var nextLayoutNum = GetNextLayoutFileNumber();
            var destLayoutPartName = $"ppt/slideLayouts/slideLayout{nextLayoutNum}.xml";

            layoutPartMapping[sourceLayoutPartName] = destLayoutPartName;

            // Clone layout XML
            ClonePartXml(sourceLayoutPartName, destLayoutPartName);

            // Clone layout rels, updating master reference to point to new master
            CloneLayoutRels(sourceLayoutPartName, destLayoutPartName, destMasterPartName);

            // Add layout relationship to new master's rels
            var relTarget = $"../slideLayouts/slideLayout{nextLayoutNum}.xml";
            destMasterRels.Add(SlideLayoutRelType, relTarget);
        }

        // Copy non-layout relationships from source master (theme, images, etc.)
        CopyNonLayoutRels(sourceRels, destMasterRels);

        // Save new master's rels
        var destMasterRelsPath = GetRelsPath(destMasterPartName);
        _package!.SetPart(destMasterRelsPath, destMasterRels.ToBytes());

        // Update master XML sldLayoutIdLst to reference new layout rIds
        UpdateMasterLayoutReferencesInternal(destMasterPartName, destMasterRels);

        // Add relationship from presentation to the new master
        var presRelTarget = $"slideMasters/slideMaster{nextMasterNum}.xml";
        var presRelId = _presentationPart!.Rels.Add(SlideMasterRelType, presRelTarget);

        // Add master reference to presentation.xml
        AddMasterReference(presRelId);
        _presentationPart.Flush();

        // Create layout slide objects for the cloned layouts
        var clonedLayouts = new List<ILayoutSlide>();
        foreach (var (_, destLayoutPartName) in layoutPartMapping)
        {
            var layout = new LayoutSlide();
            clonedLayouts.Add(layout);

            // Register in presentation's layout map
            _presentation?.LayoutSlidesMapInternal?.TryAdd(destLayoutPartName, layout);
        }

        // Load master XML and create SlidePart
        var masterData = _package.GetPart(destMasterPartName);
        SlidePart? masterPart = null;
        if (masterData is not null)
        {
            using var ms = new MemoryStream(masterData);
            var masterDoc = XDocument.Load(ms);
            masterPart = new SlidePart();
            masterPart.InitInternal(destMasterPartName);
            masterPart.Element = masterDoc.Root;
        }

        // Create MasterSlide object
        var masterSlide = new MasterSlide();
        masterSlide.PartName = destMasterPartName;
        masterSlide.PackageInternal = _package;
        masterSlide.SetPresentation(_presentation);
        if (masterPart is not null)
            masterSlide.SetMasterPart(masterPart);

        var layoutCollection = new MasterLayoutSlideCollection();
        layoutCollection.InitInternal(clonedLayouts);
        masterSlide.SetLayoutSlides(layoutCollection);

        // Register in presentation's master map
        _presentation?.MasterSlidesMapInternal?.TryAdd(destMasterPartName, masterSlide);

        Masters.Add(masterSlide);
        return masterSlide;
    }

    /// <summary>
    /// Finds the next available master slide file number.
    /// </summary>
    /// <returns>The next available number for a master slide file.</returns>
    public int GetNextMasterFileNumber()
    {
        return GetNextFileNumber("ppt/slideMasters/slideMaster");
    }

    /// <summary>
    /// Finds the next available layout slide file number.
    /// </summary>
    /// <returns>The next available number for a layout slide file.</returns>
    public int GetNextLayoutFileNumber()
    {
        return GetNextFileNumber("ppt/slideLayouts/slideLayout");
    }

    /// <summary>
    /// Clones a master slide part and its related resources (except layouts).
    /// </summary>
    /// <param name="sourcePackage">The source OPC package.</param>
    /// <param name="sourcePartName">The source master part name.</param>
    /// <param name="destPackage">The destination OPC package.</param>
    /// <param name="destPartName">The destination master part name.</param>
    /// <returns>A dictionary mapping source relationship IDs to destination relationship IDs for non-layout relationships.</returns>
    internal Dictionary<string, string> CloneMasterPart(
        OpcPackage sourcePackage, string sourcePartName,
        OpcPackage destPackage, string destPartName)
    {
        // Copy master XML
        var sourceContent = sourcePackage.GetPart(sourcePartName)
            ?? throw new InvalidOperationException($"Master slide not found: {sourcePartName}");

        destPackage.SetPart(destPartName, (byte[])sourceContent.Clone());

        // Copy relationships (theme, images, etc.) - but not layouts (handled separately)
        var sourceRels = LoadRelsFromPackage(sourcePackage, sourcePartName);
        var destRels = new RelsManager();
        var ridMapping = new Dictionary<string, string>();

        foreach (var rel in sourceRels.FindByType(SlideLayoutRelType))
        {
            // Skip layouts - they're handled separately
            continue;
        }

        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"))
        {
            var newRid = rel.IsExternal
                ? destRels.Add(rel.Type, rel.Target, isExternal: true)
                : destRels.Add(rel.Type, rel.Target);
            ridMapping[rel.Id] = newRid;
        }

        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
        {
            var newRid = rel.IsExternal
                ? destRels.Add(rel.Type, rel.Target, isExternal: true)
                : destRels.Add(rel.Type, rel.Target);
            ridMapping[rel.Id] = newRid;
        }

        // Copy any other non-layout, non-theme, non-image relationships
        foreach (var rel in GetAllRels(sourceRels))
        {
            if (rel.Type == SlideLayoutRelType)
                continue;
            if (ridMapping.ContainsKey(rel.Id))
                continue;

            var newRid = rel.IsExternal
                ? destRels.Add(rel.Type, rel.Target, isExternal: true)
                : destRels.Add(rel.Type, rel.Target);
            ridMapping[rel.Id] = newRid;
        }

        // Update r:id references in master XML
        UpdateRidReferencesInPart(destPackage, destPartName, ridMapping);

        // Save rels
        var destRelsPath = GetRelsPath(destPartName);
        destPackage.SetPart(destRelsPath, destRels.ToBytes());

        return ridMapping;
    }

    /// <summary>
    /// Clones a layout slide part.
    /// </summary>
    /// <param name="sourcePackage">The source OPC package.</param>
    /// <param name="sourcePartName">The source layout part name.</param>
    /// <param name="destPackage">The destination OPC package.</param>
    /// <param name="destPartName">The destination layout part name.</param>
    /// <param name="destMasterPartName">The destination master part name to point to.</param>
    internal void CloneLayoutPart(
        OpcPackage sourcePackage, string sourcePartName,
        OpcPackage destPackage, string destPartName,
        string destMasterPartName)
    {
        // Copy layout XML
        var sourceContent = sourcePackage.GetPart(sourcePartName)
            ?? throw new InvalidOperationException($"Layout slide not found: {sourcePartName}");

        destPackage.SetPart(destPartName, (byte[])sourceContent.Clone());

        // Copy relationships
        var sourceRels = LoadRelsFromPackage(sourcePackage, sourcePartName);
        var destRels = new RelsManager();
        var ridMapping = new Dictionary<string, string>();

        foreach (var rel in GetAllRels(sourceRels))
        {
            string newRid;
            if (rel.Type == SlideMasterRelType)
            {
                // Point to the cloned master
                var relativeTarget = GetRelativeTarget(destPartName, destMasterPartName);
                newRid = destRels.Add(rel.Type, relativeTarget);
            }
            else if (rel.IsExternal)
            {
                newRid = destRels.Add(rel.Type, rel.Target, isExternal: true);
            }
            else
            {
                newRid = destRels.Add(rel.Type, rel.Target);
            }
            ridMapping[rel.Id] = newRid;
        }

        // Update r:id references
        UpdateRidReferencesInPart(destPackage, destPartName, ridMapping);

        // Save rels
        var destRelsPath = GetRelsPath(destPartName);
        destPackage.SetPart(destRelsPath, destRels.ToBytes());
    }

    /// <summary>
    /// Finds the maximum ID across all master slide IDs and layout IDs.
    /// In PPTX, sldMasterIdLst and sldLayoutIdLst share the same ID space.
    /// </summary>
    /// <returns>The maximum ID found, or 2147483647 if none exist.</returns>
    public int GetMaxMasterLayoutIdInPresentation()
    {
        var maxId = 2147483647; // One below the PPTX convention start

        if (_presentationPart is null || _package is null)
            return maxId;

        // Check master IDs from presentation.xml
        foreach (var (masterId, _) in _presentationPart.GetMasterEntries())
        {
            if ((int)masterId > maxId)
                maxId = (int)masterId;
        }

        // Check layout IDs from all master slide XML files
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

        return maxId;
    }

    /// <summary>
    /// Updates master's relationships and XML to point to cloned layouts.
    /// </summary>
    /// <param name="sourcePackage">The source OPC package.</param>
    /// <param name="sourceMasterPartName">The source master's part name.</param>
    /// <param name="destMasterPartName">The destination master's part name.</param>
    /// <param name="layoutMapping">Dictionary of old layout path to new layout path.</param>
    /// <param name="sourceLayoutRids">Dictionary of source layout path to source relationship ID.</param>
    internal void UpdateMasterLayoutRelationships(
        OpcPackage sourcePackage, string sourceMasterPartName,
        string destMasterPartName,
        Dictionary<string, string> layoutMapping,
        Dictionary<string, string> sourceLayoutRids)
    {
        if (_package is null)
            return;

        var masterRels = LoadRelsFromPackage(_package, destMasterPartName);
        var layoutRidMapping = new Dictionary<string, string>();

        foreach (var (oldLayoutPath, newLayoutPath) in layoutMapping)
        {
            var relativeTarget = GetRelativeTarget(destMasterPartName, newLayoutPath);
            var newRid = masterRels.Add(SlideLayoutRelType, relativeTarget);

            // Map old rid to new rid
            if (sourceLayoutRids.TryGetValue(oldLayoutPath, out var oldRid))
            {
                layoutRidMapping[oldRid] = newRid;
            }
        }

        // Save updated rels
        var relsPath = GetRelsPath(destMasterPartName);
        _package.SetPart(relsPath, masterRels.ToBytes());

        // Update the master XML's sldLayoutIdLst references
        var masterContent = _package.GetPart(destMasterPartName);
        if (masterContent is not null)
        {
            using var msMaster = new MemoryStream(masterContent);
            var doc = XDocument.Load(msMaster);
            var root = doc.Root;
            if (root is not null)
            {
                // Update r:id references in sldLayoutIdLst
                UpdateRidReferences(root, layoutRidMapping);

                // Renumber layout IDs to avoid conflicts with existing masters
                var nextLayoutId = GetMaxMasterLayoutIdInPresentation() + 1;
                foreach (var elem in root.Descendants(PNs + "sldLayoutId"))
                {
                    elem.SetAttributeValue("id", nextLayoutId.ToString());
                    nextLayoutId++;
                }

                // Save updated master XML
                using var ms = new MemoryStream();
                doc.Save(ms);
                _package.SetPart(destMasterPartName, ms.ToArray());
            }
        }
    }

    // ── Private helpers ──────────────────────────────────────────

    private int GetNextFileNumber(string pathPrefix)
    {
        var existing = new HashSet<int>();
        var pattern = new Regex(Regex.Escape(pathPrefix) + @"(\d+)\.xml$", RegexOptions.IgnoreCase);
        foreach (var partName in _package!.GetPartNames())
        {
            var m = pattern.Match(partName);
            if (m.Success)
                existing.Add(int.Parse(m.Groups[1].Value));
        }

        var num = 1;
        while (existing.Contains(num))
            num++;
        return num;
    }

    private void ClonePartXml(string sourcePartName, string destPartName)
    {
        var sourceData = _package!.GetPart(sourcePartName);
        if (sourceData is not null)
        {
            _package.SetPart(destPartName, (byte[])sourceData.Clone());
        }
    }

    private RelsManager LoadRels(string partName)
    {
        var rels = new RelsManager();
        var relsPath = GetRelsPath(partName);
        var relsData = _package!.GetPart(relsPath);
        if (relsData is not null)
            rels.Load(relsData);
        return rels;
    }

    private static RelsManager LoadRelsFromPackage(OpcPackage package, string partName)
    {
        var rels = new RelsManager();
        var relsPath = GetRelsPath(partName);
        var relsData = package.GetPart(relsPath);
        if (relsData is not null)
            rels.Load(relsData);
        return rels;
    }

    private static List<Relationship> GetAllRels(RelsManager rels)
    {
        // Collect all relationships by iterating known types and any others
        var result = new List<Relationship>();
        var seen = new HashSet<string>();
        foreach (var rel in rels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"))
        {
            if (seen.Add(rel.Id)) result.Add(rel);
        }
        foreach (var rel in rels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"))
        {
            if (seen.Add(rel.Id)) result.Add(rel);
        }
        foreach (var rel in rels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"))
        {
            if (seen.Add(rel.Id)) result.Add(rel);
        }
        foreach (var rel in rels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
        {
            if (seen.Add(rel.Id)) result.Add(rel);
        }
        foreach (var rel in rels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
        {
            if (seen.Add(rel.Id)) result.Add(rel);
        }
        return result;
    }

    private void CopyNonLayoutRels(RelsManager sourceRels, RelsManager destRels)
    {
        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"))
        {
            destRels.Add(rel.Type, rel.Target);
        }

        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
        {
            destRels.Add(rel.Type, rel.Target);
        }
    }

    private void CloneLayoutRels(string sourceLayoutPartName, string destLayoutPartName,
        string destMasterPartName)
    {
        var sourceRels = LoadRels(sourceLayoutPartName);
        var destRels = new RelsManager();

        // Add relationship to the new master
        var masterRelTarget = GetRelativeTarget(destLayoutPartName, destMasterPartName);
        destRels.Add(SlideMasterRelType, masterRelTarget);

        // Copy non-master relationships
        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"))
        {
            destRels.Add(rel.Type, rel.Target);
        }

        foreach (var rel in sourceRels.FindByType(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
        {
            destRels.Add(rel.Type, rel.Target);
        }

        var destRelsPath = GetRelsPath(destLayoutPartName);
        _package!.SetPart(destRelsPath, destRels.ToBytes());
    }

    private void UpdateMasterLayoutReferencesInternal(string masterPartName, RelsManager masterRels)
    {
        var data = _package!.GetPart(masterPartName);
        if (data is null)
            return;

        using var msData = new MemoryStream(data);
        var doc = XDocument.Load(msData);
        var root = doc.Root;
        if (root is null)
            return;

        // Remove existing sldLayoutIdLst
        root.Element(PNs + "sldLayoutIdLst")?.Remove();

        // Build new sldLayoutIdLst with the cloned layout rIds
        var layoutIdLst = new XElement(PNs + "sldLayoutIdLst");
        uint layoutBaseId = 2147483649;

        foreach (var rel in masterRels.FindByType(SlideLayoutRelType))
        {
            layoutIdLst.Add(new XElement(PNs + "sldLayoutId",
                new XAttribute("id", layoutBaseId++),
                new XAttribute(RNs + "id", rel.Id)));
        }

        // Insert after clrMap or at end
        var clrMap = root.Element(PNs + "clrMap");
        if (clrMap is not null)
            clrMap.AddAfterSelf(layoutIdLst);
        else
            root.Add(layoutIdLst);

        // Save updated XML back
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(masterPartName, ms.ToArray());
    }

    private void AddMasterReference(string relId)
    {
        var root = _presentationPart!.Document.Root;
        if (root is null)
            return;

        var masterIdLst = root.Element(PNs + "sldMasterIdLst");
        if (masterIdLst is null)
        {
            masterIdLst = new XElement(PNs + "sldMasterIdLst");
            root.AddFirst(masterIdLst);
        }

        // Find next master ID
        uint nextId = 2147483648;
        foreach (var existing in masterIdLst.Elements(PNs + "sldMasterId"))
        {
            var idAttr = existing.Attribute("id");
            if (idAttr is not null && uint.TryParse(idAttr.Value, out var id) && id >= nextId)
                nextId = id + 1;
        }

        masterIdLst.Add(new XElement(PNs + "sldMasterId",
            new XAttribute("id", nextId),
            new XAttribute(RNs + "id", relId)));
    }

    private static void UpdateRidReferences(XElement root, Dictionary<string, string> ridMapping)
    {
        if (ridMapping.Count == 0)
            return;

        foreach (var elem in root.DescendantsAndSelf())
        {
            foreach (var attr in elem.Attributes())
            {
                if (ridMapping.TryGetValue(attr.Value, out var newRid))
                {
                    attr.Value = newRid;
                }
            }
        }
    }

    private static void UpdateRidReferencesInPart(
        OpcPackage package, string partName, Dictionary<string, string> ridMapping)
    {
        if (ridMapping.Count == 0)
            return;

        var content = package.GetPart(partName);
        if (content is null)
            return;

        using var msContent = new MemoryStream(content);
        var doc = XDocument.Load(msContent);
        if (doc.Root is null)
            return;

        UpdateRidReferences(doc.Root, ridMapping);

        using var ms = new MemoryStream();
        doc.Save(ms);
        package.SetPart(partName, ms.ToArray());
    }

    private static string GetRelativeTarget(string fromPartName, string toPartName)
    {
        var fromDir = GetDirectoryPath(fromPartName);
        var toDir = GetDirectoryPath(toPartName);
        var toFile = toPartName[toDir.Length..];

        if (fromDir == toDir)
            return toFile;

        // Simple one-level relative path
        var fromSegments = fromDir.TrimEnd('/').Split('/');
        var toSegments = toDir.TrimEnd('/').Split('/');

        // Find common prefix
        var common = 0;
        while (common < fromSegments.Length && common < toSegments.Length &&
               string.Equals(fromSegments[common], toSegments[common], StringComparison.OrdinalIgnoreCase))
        {
            common++;
        }

        var ups = fromSegments.Length - common;
        var relative = string.Join("/", Enumerable.Repeat("..", ups));
        var remainingPath = string.Join("/", toSegments.Skip(common));
        if (remainingPath.Length > 0)
            relative = relative + "/" + remainingPath;

        return relative + "/" + toFile;
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
}
