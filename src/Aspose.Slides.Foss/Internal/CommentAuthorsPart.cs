using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages the comment authors XML part (ppt/commentAuthors.xml).
/// </summary>
internal sealed class CommentAuthorsPart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private readonly XElement _root;

    internal CommentAuthorsPart(XElement root)
    {
        _root = root;
    }

    internal AuthorData? FindAuthorById(int authorId)
    {
        var elem = _root.Elements(PNs + "cmAuthor")
            .FirstOrDefault(e => int.Parse(e.Attribute("id")?.Value ?? "-1") == authorId);
        return elem is not null ? new AuthorData(elem) : null;
    }

    /// <summary>
    /// Returns all author data records from the XML.
    /// </summary>
    internal List<AuthorData> GetAuthors()
    {
        return _root.Elements(PNs + "cmAuthor")
            .Select(e => new AuthorData(e))
            .ToList();
    }

    /// <summary>
    /// Adds a new author element and returns its data.
    /// </summary>
    internal AuthorData AddAuthor(string name, string initials)
    {
        var existingIds = _root.Elements(PNs + "cmAuthor")
            .Select(e => int.Parse(e.Attribute("id")?.Value ?? "0"))
            .ToList();

        var newId = existingIds.Count > 0 ? existingIds.Max() + 1 : 0;

        var elem = new XElement(PNs + "cmAuthor",
            new XAttribute("id", newId),
            new XAttribute("name", name),
            new XAttribute("initials", initials),
            new XAttribute("lastIdx", 0),
            new XAttribute("clrIdx", newId));

        _root.Add(elem);

        return new AuthorData(elem);
    }

    /// <summary>
    /// Removes the author element with the specified ID from the XML.
    /// </summary>
    internal void RemoveAuthor(int authorId)
    {
        var elem = _root.Elements(PNs + "cmAuthor")
            .FirstOrDefault(e => int.Parse(e.Attribute("id")?.Value ?? "-1") == authorId);
        elem?.Remove();
    }

    /// <summary>
    /// Removes all author elements from the XML.
    /// </summary>
    internal void Clear()
    {
        _root.Elements(PNs + "cmAuthor").Remove();
    }

    /// <summary>
    /// Gets a shared cache of <see cref="CommentsPart"/> instances keyed by slide part name.
    /// Used by <see cref="Aspose.Slides.Foss.CommentCollection"/> to avoid redundant loads during writes.
    /// </summary>
    internal Dictionary<string, CommentsPart> CpCache { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Allocates and returns the next unique comment index for the specified author.
    /// Increments the author's <c>lastIdx</c> attribute and returns the new value.
    /// </summary>
    /// <param name="authorId">The ID of the author.</param>
    /// <returns>The next comment index.</returns>
    internal int NextCommentIdx(int authorId)
    {
        var elem = _root.Elements(PNs + "cmAuthor")
            .FirstOrDefault(e => int.Parse(e.Attribute("id")?.Value ?? "-1") == authorId);

        if (elem is null)
            return 1;

        var lastIdx = int.Parse(elem.Attribute("lastIdx")?.Value ?? "0");
        var newIdx = lastIdx + 1;
        elem.SetAttributeValue("lastIdx", newIdx);
        return newIdx;
    }

    /// <summary>
    /// Persists changes to the authors part.
    /// </summary>
    internal void SaveAuthors()
    {
        // Placeholder for the full OPC save logic.
    }

    /// <summary>
    /// Flushes the authors XML and all cached comment parts back to the OPC package.
    /// </summary>
    internal void Flush(OpcPackage package)
    {
        const string authorsPartName = "ppt/commentAuthors.xml";

        // Write commentAuthors.xml
        var authorsDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), _root);
        using (var ms = new MemoryStream())
        {
            authorsDoc.Save(ms);
            package.SetPart(authorsPartName, ms.ToArray());
        }

        // Write each cached comments part
        foreach (var (partName, cp) in CpCache)
        {
            cp.Flush(package, partName);
        }
    }
}
