using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CommentAuthorsPart;

/// <summary>
/// Manages the comment authors XML part (<c>ppt/commentAuthors.xml</c>).
/// This part holds all author definitions used across the presentation's comments.
/// </summary>
public sealed class CommentAuthorsPart
{
    private static readonly XNamespace PNs =
        "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static readonly XNamespace CtNs =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    private const string CommentAuthorsContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";

    private const string CommentAuthorsRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";

    /// <summary>
    /// The OPC part name for the comment authors part.
    /// </summary>
    public const string PartName = "ppt/commentAuthors.xml";

    private readonly OpcPackage _package;
    private XElement _root;

    /// <summary>
    /// Initializes a new <see cref="CommentAuthorsPart"/> and loads existing data from the package.
    /// </summary>
    internal CommentAuthorsPart(OpcPackage package)
    {
        _package = package;
        _root = null!;
        Load();
    }

    /// <summary>
    /// Loads the comment authors XML from the package, or creates an empty root element.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(PartName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            var doc = XDocument.Load(ms);
            _root = doc.Root ?? new XElement(PNs + "cmAuthorLst",
                new XAttribute(XNamespace.Xmlns + "p", PNs.NamespaceName));
        }
        else
        {
            _root = new XElement(PNs + "cmAuthorLst",
                new XAttribute(XNamespace.Xmlns + "p", PNs.NamespaceName));
        }
    }

    /// <summary>
    /// Returns all comment authors.
    /// </summary>
    public List<AuthorData> GetAuthors()
    {
        return _root.Elements(PNs + "cmAuthor").Select(e => new AuthorData(e)).ToList();
    }

    /// <summary>
    /// Finds an author by their ID.
    /// </summary>
    /// <param name="authorId">The author ID to search for.</param>
    /// <returns>The <see cref="AuthorData"/> or <c>null</c> if not found.</returns>
    public AuthorData? FindAuthorById(int authorId)
    {
        var elem = _root.Elements(PNs + "cmAuthor")
            .FirstOrDefault(e => int.Parse(e.Attribute("id")?.Value ?? "-1") == authorId);
        return elem is not null ? new AuthorData(elem) : null;
    }

    /// <summary>
    /// Adds a new author and returns its <see cref="AuthorData"/>.
    /// </summary>
    /// <param name="name">The author's display name.</param>
    /// <param name="initials">The author's initials.</param>
    public AuthorData AddAuthor(string name, string initials)
    {
        var existing = _root.Elements(PNs + "cmAuthor").ToList();
        var nextId = existing.Count;
        var clrIdx = nextId % 10;

        var elem = new XElement(PNs + "cmAuthor",
            new XAttribute("id", nextId),
            new XAttribute("name", name),
            new XAttribute("initials", initials),
            new XAttribute("lastIdx", 0),
            new XAttribute("clrIdx", clrIdx));

        _root.Add(elem);
        return new AuthorData(elem);
    }

    /// <summary>
    /// Removes the author with the specified ID.
    /// </summary>
    /// <param name="authorId">The author ID to remove.</param>
    public void RemoveAuthor(int authorId)
    {
        var elem = _root.Elements(PNs + "cmAuthor")
            .FirstOrDefault(e => int.Parse(e.Attribute("id")?.Value ?? "-1") == authorId);
        elem?.Remove();
    }

    /// <summary>
    /// Removes all author elements.
    /// </summary>
    public void Clear()
    {
        _root.Elements(PNs + "cmAuthor").Remove();
    }

    /// <summary>
    /// Returns the next globally-unique comment index.
    /// OOXML parentCmId references idx values that must be unique across ALL
    /// authors within the presentation (not just per-author). This method takes
    /// the maximum lastIdx across every author, increments it, and updates
    /// the target author's lastIdx.
    /// </summary>
    /// <param name="authorId">The author ID to update.</param>
    /// <returns>The new globally-unique comment index.</returns>
    public int NextCommentIdx(int authorId)
    {
        var globalMax = 0;
        foreach (var e in _root.Elements(PNs + "cmAuthor"))
        {
            globalMax = Math.Max(globalMax, int.Parse(e.Attribute("lastIdx")?.Value ?? "0"));
        }

        var newIdx = globalMax + 1;
        var data = FindAuthorById(authorId);
        if (data is not null)
        {
            data.LastIdx = newIdx;
        }

        return newIdx;
    }

    /// <summary>
    /// Saves the comment authors XML back to the package.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());
    }

    /// <summary>
    /// Ensures that commentAuthors.xml is registered in content types and
    /// the presentation has a relationship to it. Call before first save.
    /// </summary>
    internal static void EnsureRegistered(OpcPackage package)
    {
        // Add content type override if needed
        AddContentTypeOverride(package, PartName, CommentAuthorsContentType);

        // Add relationship from presentation to commentAuthors if needed
        var rels = LoadRelsManager(package, PresentationPart.PresentationPart.PartName);
        var existing = rels.FindByType(CommentAuthorsRelType).ToList();
        if (existing.Count == 0)
        {
            rels.Add(CommentAuthorsRelType, "commentAuthors.xml");
            SaveRelsManager(package, PresentationPart.PresentationPart.PartName, rels);
        }
    }

    #region Private helpers

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
        else if (existing.Attribute("ContentType")?.Value != contentType)
        {
            existing.SetAttributeValue("ContentType", contentType);
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    #endregion
}
