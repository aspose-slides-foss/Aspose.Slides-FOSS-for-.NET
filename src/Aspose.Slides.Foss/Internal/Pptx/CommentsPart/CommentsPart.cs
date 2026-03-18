using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CommentsPart;

/// <summary>
/// Manages a slide comments XML part (<c>ppt/comments/slideN.xml</c>).
/// One file exists per slide that has comments.
/// </summary>
public sealed class CommentsPart
{
    private const int CmToEmu = 360000;

    private static readonly XNamespace PNs =
        "http://schemas.openxmlformats.org/presentationml/2006/main";

    private const string CommentsRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    private const string CommentsContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";

    private readonly OpcPackage _package;
    private readonly string _partName;
    private XElement _root;

    /// <summary>
    /// Initializes a <see cref="CommentsPart"/> from an existing part in the package.
    /// </summary>
    internal CommentsPart(OpcPackage package, string partName)
    {
        _package = package;
        _partName = partName;
        _root = null!;
        Load();
    }

    // Internal constructor for CreateForSlide (root already built).
    private CommentsPart(OpcPackage package, string partName, XElement root)
    {
        _package = package;
        _partName = partName;
        _root = root;
    }

    /// <summary>
    /// Gets the part name of this comments part.
    /// </summary>
    public string PartName => _partName;

    /// <summary>
    /// Loads the comments XML from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(_partName);
        if (content is null)
            throw new InvalidOperationException($"Comments part not found: {_partName}");

        using var ms = new MemoryStream(content);
        var doc = XDocument.Load(ms);
        _root = doc.Root ?? throw new InvalidOperationException($"Invalid comments XML in {_partName}");
    }

    /// <summary>
    /// Returns all comments in this part.
    /// </summary>
    public List<CommentData> GetComments()
    {
        return _root.Elements(PNs + "cm").Select(e => new CommentData(e)).ToList();
    }

    /// <summary>
    /// Returns comments belonging to the specified author.
    /// </summary>
    public List<CommentData> GetCommentsByAuthor(int authorId)
    {
        return _root.Elements(PNs + "cm")
            .Where(e => int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId)
            .Select(e => new CommentData(e))
            .ToList();
    }

    /// <summary>
    /// Finds a comment by its index across all authors.
    /// </summary>
    public CommentData? FindCommentByIdxAll(int idx)
    {
        var elem = _root.Elements(PNs + "cm")
            .FirstOrDefault(e => int.Parse(e.Attribute("idx")?.Value ?? "-1") == idx);
        return elem is not null ? new CommentData(elem) : null;
    }

    /// <summary>
    /// Finds a comment by author ID and index.
    /// </summary>
    public CommentData? FindCommentByIdx(int authorId, int idx)
    {
        var elem = _root.Elements(PNs + "cm")
            .FirstOrDefault(e =>
                int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId &&
                int.Parse(e.Attribute("idx")?.Value ?? "-1") == idx);
        return elem is not null ? new CommentData(elem) : null;
    }

    /// <summary>
    /// Appends a new comment element and returns its data.
    /// </summary>
    public CommentData AddComment(int authorId, int idx, string text, float posX, float posY,
        string dtStr, int? parentIdx = null)
    {
        var elem = CreateCommentElement(authorId, idx, text, posX, posY, dtStr, parentIdx);
        _root.Add(elem);
        return new CommentData(elem);
    }

    /// <summary>
    /// Inserts a comment at the given index among existing comments.
    /// </summary>
    public CommentData InsertComment(int index, int authorId, int idx, string text, float posX,
        float posY, string dtStr, int? parentIdx = null)
    {
        var elem = CreateCommentElement(authorId, idx, text, posX, posY, dtStr, parentIdx);
        var existing = _root.Elements(PNs + "cm").ToList();
        if (index >= 0 && index < existing.Count)
            existing[index].AddBeforeSelf(elem);
        else
            _root.Add(elem);
        return new CommentData(elem);
    }

    /// <summary>
    /// Removes a comment by author ID and index.
    /// </summary>
    public void RemoveComment(int authorId, int idx)
    {
        var elem = _root.Elements(PNs + "cm")
            .FirstOrDefault(e =>
                int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId &&
                int.Parse(e.Attribute("idx")?.Value ?? "-1") == idx);
        elem?.Remove();
    }

    /// <summary>
    /// Removes the specified XML element from the comments list.
    /// </summary>
    public void RemoveCommentElem(XElement elem)
    {
        try { elem.Remove(); }
        catch (InvalidOperationException) { /* already removed */ }
    }

    /// <summary>
    /// Removes the comment at the specified position.
    /// </summary>
    public void RemoveCommentsAt(int index)
    {
        var all = _root.Elements(PNs + "cm").ToList();
        if (index >= 0 && index < all.Count)
            all[index].Remove();
    }

    /// <summary>
    /// Removes all comment elements.
    /// </summary>
    public void Clear()
    {
        _root.Elements(PNs + "cm").Remove();
    }

    /// <summary>
    /// Returns the number of comments.
    /// </summary>
    public int Count()
    {
        return _root.Elements(PNs + "cm").Count();
    }

    /// <summary>
    /// Returns <c>true</c> if there are no comments.
    /// </summary>
    public bool IsEmpty()
    {
        return Count() == 0;
    }

    /// <summary>
    /// Saves the comments XML back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(_partName, ms.ToArray());
    }

    /// <summary>
    /// Creates a new empty comments part for a slide and registers all relationships.
    /// </summary>
    /// <param name="package">The OPC package.</param>
    /// <param name="slidePartName">The slide part name (e.g. <c>ppt/slides/slide1.xml</c>).</param>
    /// <param name="slideRelsManager">
    /// If provided, the slide's existing <see cref="RelsManager"/> is used so the relationship
    /// survives subsequent saves. If <c>null</c>, a fresh one is created.
    /// </param>
    internal static CommentsPart CreateForSlide(OpcPackage package, string slidePartName,
        RelsManager? slideRelsManager = null)
    {
        // Find a unique part name
        int num = 1;
        string partName;
        while (true)
        {
            var candidate = $"ppt/comments/slide{num}.xml";
            if (package.GetPart(candidate) is null)
            {
                partName = candidate;
                break;
            }
            num++;
        }

        // Build minimal XML
        var root = new XElement(PNs + "cmLst",
            new XAttribute(XNamespace.Xmlns + "p", PNs.NamespaceName));

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        using (var ms = new MemoryStream())
        {
            doc.Save(ms);
            package.SetPart(partName, ms.ToArray());
        }

        // Content type override
        AddContentTypeOverride(package, partName, CommentsContentType);

        // Relationship: slide → comments
        var relTarget = ComputeRelativeTarget(slidePartName, partName);
        var rels = slideRelsManager ?? LoadRelsManager(package, slidePartName);
        rels.Add(CommentsRelType, relTarget);
        SaveRelsManager(package, slidePartName, rels);

        return new CommentsPart(package, partName, root);
    }

    /// <summary>
    /// Loads the comments part for a slide, if it exists.
    /// </summary>
    internal static CommentsPart? LoadForSlide(OpcPackage package, string slidePartName)
    {
        var rels = LoadRelsManager(package, slidePartName);
        var commentRels = rels.FindByType(CommentsRelType).ToList();
        if (commentRels.Count == 0)
            return null;

        var target = commentRels[0].Target;
        var partName = ResolveTarget(slidePartName, target);
        return package.GetPart(partName) is not null
            ? new CommentsPart(package, partName)
            : null;
    }

    /// <summary>
    /// Deletes the comments part for a slide.
    /// </summary>
    internal static void Delete(OpcPackage package, string slidePartName)
    {
        var rels = LoadRelsManager(package, slidePartName);
        var commentRels = rels.FindByType(CommentsRelType).ToList();
        if (commentRels.Count == 0)
            return;

        foreach (var rel in commentRels)
        {
            var partName = ResolveTarget(slidePartName, rel.Target);
            package.RemovePart(partName);
            RemoveContentTypeOverride(package, partName);
            rels.Remove(rel.Id);
        }

        SaveRelsManager(package, slidePartName, rels);
    }

    /// <summary>
    /// Resolves a relative target path to an absolute part name.
    /// </summary>
    public static string ResolveTarget(string fromPart, string target)
    {
        if (target.StartsWith('/'))
            return target.TrimStart('/');

        var baseDir = fromPart.Contains('/')
            ? fromPart[..fromPart.LastIndexOf('/')]
            : "";

        var parts = (baseDir + "/" + target).Split('/');
        var resolved = new List<string>();
        foreach (var part in parts)
        {
            if (part == "..")
            {
                if (resolved.Count > 0) resolved.RemoveAt(resolved.Count - 1);
            }
            else if (part is not "" and not ".")
            {
                resolved.Add(part);
            }
        }
        return string.Join("/", resolved);
    }

    /// <summary>
    /// Computes a relative path from one part to another.
    /// </summary>
    public static string ComputeRelativeTarget(string fromPart, string toPart)
    {
        var fromDir = fromPart.Contains('/')
            ? fromPart[..fromPart.LastIndexOf('/')]
            : "";
        var toDir = toPart.Contains('/')
            ? toPart[..toPart.LastIndexOf('/')]
            : "";
        var toFile = toPart.Contains('/')
            ? toPart[(toPart.LastIndexOf('/') + 1)..]
            : toPart;

        if (fromDir == toDir)
            return toFile;

        var fromParts = fromDir.Length > 0 ? fromDir.Split('/') : [];
        var toParts = toDir.Length > 0 ? toDir.Split('/') : [];

        int commonLen = 0;
        for (int i = 0; i < Math.Min(fromParts.Length, toParts.Length); i++)
        {
            if (fromParts[i] == toParts[i])
                commonLen = i + 1;
            else
                break;
        }

        int upCount = fromParts.Length - commonLen;
        var downPath = string.Join("/", toParts.Skip(commonLen));

        var result = string.Concat(Enumerable.Repeat("../", upCount));
        if (downPath.Length > 0)
            result += downPath + "/";
        result += toFile;
        return result;
    }

    #region Private helpers

    private static XElement CreateCommentElement(int authorId, int idx, string text,
        float posX, float posY, string dtStr, int? parentIdx)
    {
        var elem = new XElement(PNs + "cm",
            new XAttribute("authorId", authorId),
            new XAttribute("dt", dtStr),
            new XAttribute("idx", idx));

        if (parentIdx is not null)
            elem.Add(new XAttribute("parentCmId", parentIdx.Value));

        elem.Add(new XElement(PNs + "pos",
            new XAttribute("x", (int)Math.Round(posX * CmToEmu)),
            new XAttribute("y", (int)Math.Round(posY * CmToEmu))));

        elem.Add(new XElement(PNs + "text", text));

        return elem;
    }

    private static string GetRelsPartName(string partName)
    {
        if (partName.Contains('/'))
        {
            var dir = partName[..partName.LastIndexOf('/')];
            var file = partName[(partName.LastIndexOf('/') + 1)..];
            return $"{dir}/_rels/{file}.rels";
        }
        return $"_rels/{partName}.rels";
    }

    private static RelsManager LoadRelsManager(OpcPackage package, string partName)
    {
        var relsPath = GetRelsPartName(partName);
        var rels = new RelsManager();
        var data = package.GetPart(relsPath);
        if (data is not null)
            rels.Load(data);
        return rels;
    }

    private static void SaveRelsManager(OpcPackage package, string partName, RelsManager rels)
    {
        var relsPath = GetRelsPartName(partName);
        package.SetPart(relsPath, rels.ToBytes());
    }

    private static readonly XNamespace CtNs =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    private static void AddContentTypeOverride(OpcPackage package, string partName, string contentType)
    {
        var ctData = package.GetPart("[Content_Types].xml");
        XDocument doc;
        if (ctData is not null)
        {
            using var ms = new MemoryStream(ctData);
            doc = XDocument.Load(ms);
        }
        else
        {
            doc = new XDocument(new XElement(CtNs + "Types"));
        }

        var root = doc.Root ?? throw new InvalidOperationException("Content types XML has no root element");
        // Check if override already exists
        var existing = root.Elements(CtNs + "Override")
            .FirstOrDefault(e => string.Equals(
                e.Attribute("PartName")?.Value?.TrimStart('/'),
                partName.TrimStart('/'),
                StringComparison.OrdinalIgnoreCase));

        if (existing is null)
        {
            root.Add(new XElement(CtNs + "Override",
                new XAttribute("PartName", "/" + partName),
                new XAttribute("ContentType", contentType)));
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    private static void RemoveContentTypeOverride(OpcPackage package, string partName)
    {
        var ctData = package.GetPart("[Content_Types].xml");
        if (ctData is null) return;

        using var ms = new MemoryStream(ctData);
        var doc = XDocument.Load(ms);
        var root = doc.Root ?? throw new InvalidOperationException("Content types XML has no root element");

        var toRemove = root.Elements(CtNs + "Override")
            .Where(e => string.Equals(
                e.Attribute("PartName")?.Value?.TrimStart('/'),
                partName.TrimStart('/'),
                StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var elem in toRemove)
            elem.Remove();

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    #endregion
}
