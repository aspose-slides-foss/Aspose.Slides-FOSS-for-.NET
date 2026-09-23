using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages a slide comments XML part (ppt/comments/slideN.xml).
/// </summary>
internal sealed class CommentsPart
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private readonly XElement _root;

    internal CommentsPart(XElement root)
    {
        _root = root;
    }

    internal List<CommentData> GetComments()
    {
        return _root.Elements(PNs + "cm").Select(e => new CommentData(e)).ToList();
    }

    internal List<CommentData> GetCommentsByAuthor(int authorId)
    {
        return _root.Elements(PNs + "cm")
            .Where(e => int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId)
            .Select(e => new CommentData(e))
            .ToList();
    }

    /// <summary>
    /// Finds a comment by its author and index. An index is counted per author, so two authors'
    /// comments on one slide can share it.
    /// </summary>
    internal CommentData? FindComment(int authorId, int idx)
    {
        return GetComments().FirstOrDefault(cd => cd.AuthorId == authorId && cd.Idx == idx);
    }

    /// <summary>
    /// Returns the comments in this part that reply to the given comment.
    /// </summary>
    internal List<CommentData> GetReplies(CommentData parent)
    {
        var parentRef = (parent.AuthorId, parent.Idx);
        return GetComments().Where(cd => cd.ParentRef == parentRef).ToList();
    }

    internal void RemoveCommentElem(XElement elem)
    {
        elem.Remove();
    }

    /// <summary>
    /// Appends a new comment element and returns its data.
    /// </summary>
    internal CommentData AddComment(int authorId, int idx, string text, float posX, float posY, string dtStr)
    {
        var elem = CreateCommentElement(authorId, idx, text, posX, posY, dtStr);
        _root.Add(elem);
        return new CommentData(elem);
    }

    /// <summary>
    /// Inserts a new comment element at the specified position and returns its data.
    /// </summary>
    internal CommentData InsertComment(int index, int authorId, int idx, string text, float posX, float posY, string dtStr)
    {
        var elem = CreateCommentElement(authorId, idx, text, posX, posY, dtStr);
        var existing = _root.Elements(PNs + "cm").ToList();
        if (index >= 0 && index < existing.Count)
            existing[index].AddBeforeSelf(elem);
        else
            _root.Add(elem);
        return new CommentData(elem);
    }

    /// <summary>
    /// Removes all comment elements belonging to the specified author.
    /// </summary>
    internal void RemoveCommentsByAuthor(int authorId)
    {
        var toRemove = _root.Elements(PNs + "cm")
            .Where(e => int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId)
            .ToList();
        foreach (var elem in toRemove)
            elem.Remove();
    }

    /// <summary>
    /// Creates a new empty <see cref="CommentsPart"/> with an empty root element.
    /// </summary>
    internal static CommentsPart CreateEmpty()
    {
        var root = new XElement(PNs + "cmLst");
        return new CommentsPart(root);
    }

    /// <summary>
    /// Loads a <see cref="CommentsPart"/> from the package for the given part name,
    /// or returns <c>null</c> if no such part exists.
    /// </summary>
    internal static CommentsPart? LoadFromPackage(OpcPackage package, string commentsPartName)
    {
        var data = package.GetPart(commentsPartName);
        if (data is null)
            return null;

        using var stream = new System.IO.MemoryStream(data);
        var doc = XDocument.Load(stream);
        return doc.Root is not null ? new CommentsPart(doc.Root) : null;
    }

    internal void Save()
    {
        // Persist changes - in real implementation this writes back to the package.
        // Placeholder for the full OPC save logic.
    }

    /// <summary>
    /// Flushes the comments XML back to the OPC package at the given part name.
    /// </summary>
    /// <remarks>
    /// A reply names its parent by author and index. A reply whose parent is no longer in the part —
    /// removed on its own, or with its author — is written as a top-level comment rather than as a
    /// reference to a comment the file does not have.
    /// </remarks>
    internal void Flush(OpcPackage package, string partName)
    {
        var serialized = new XElement(_root);
        var comments = serialized.Elements(PNs + "cm").Select(e => new CommentData(e)).ToList();
        var present = comments.Select(cd => (cd.AuthorId, cd.Idx)).ToHashSet();
        foreach (var cd in comments)
        {
            if (cd.ParentRef is { } parent && !present.Contains(parent))
                cd.ParentRef = null;
        }

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), serialized);
        using var ms = new MemoryStream();
        doc.Save(ms);
        package.SetPart(partName, ms.ToArray());
    }

    private static XElement CreateCommentElement(int authorId, int idx, string text, float posX, float posY, string dtStr)
    {
        // posX and posY are centimetres; p:pos is in PowerPoint's unit, not EMU (see CommentData).
        return new XElement(PNs + "cm",
            new XAttribute("authorId", authorId),
            new XAttribute("idx", idx),
            new XAttribute("dt", dtStr),
            new XElement(PNs + "pos",
                new XAttribute("x", CommentData.ToPositionUnits(posX)),
                new XAttribute("y", CommentData.ToPositionUnits(posY))),
            new XElement(PNs + "text", text));
    }
}
