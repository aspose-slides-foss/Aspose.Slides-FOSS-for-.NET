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
    /// Returns whether the comment is still in this part, that is, has not been removed.
    /// </summary>
    internal bool Contains(CommentData cd) => ReferenceEquals(cd.Elem.Parent, _root);

    /// <summary>
    /// Returns the comment the given one replies to, or <c>null</c> if it is top-level or its
    /// parent is not in this part.
    /// </summary>
    internal CommentData? FindParent(CommentData cd) =>
        cd.ParentRef is { } parentRef ? FindComment(parentRef.AuthorId, parentRef.Idx) : null;

    /// <summary>
    /// Returns the comments in this part that reply to the given comment.
    /// </summary>
    internal List<CommentData> GetReplies(CommentData parent)
    {
        var parentRef = (parent.AuthorId, parent.Idx);
        return GetComments().Where(cd => cd.ParentRef == parentRef).ToList();
    }

    /// <summary>
    /// Returns every reply under the given comment, at any depth, not including the comment itself.
    /// </summary>
    internal List<CommentData> GetDescendants(CommentData root)
    {
        var seen = new HashSet<XElement> { root.Elem };
        var result = new List<CommentData>();
        var pending = new Queue<CommentData>();
        pending.Enqueue(root);
        while (pending.Count > 0)
        {
            foreach (var reply in GetReplies(pending.Dequeue()))
            {
                if (!seen.Add(reply.Elem))
                    continue;
                result.Add(reply);
                pending.Enqueue(reply);
            }
        }
        return result;
    }

    internal void RemoveCommentElem(XElement elem) => RemoveCommentElems([elem]);

    /// <summary>
    /// Removes the given comment elements, then detaches every reply left naming a comment that is
    /// no longer here.
    /// </summary>
    internal void RemoveCommentElems(IEnumerable<XElement> elems)
    {
        foreach (var elem in elems.ToList())
            elem.Remove();
        DetachOrphanedReplies();
    }

    /// <summary>
    /// Makes every reply whose parent is not in this part a top-level comment.
    /// </summary>
    /// <remarks>
    /// A reply names its parent by author id and index only. Author ids are reused — a new author
    /// takes the highest id plus one, and so the id of a removed author — and a new author's indices
    /// start again from 1, so a reference left naming a removed comment would later name an
    /// unrelated one. It is therefore dropped as soon as the comment is gone, and when a part is
    /// loaded, rather than when the part is saved.
    /// </remarks>
    internal void DetachOrphanedReplies()
    {
        var comments = GetComments();
        var present = comments.Select(cd => (cd.AuthorId, cd.Idx)).ToHashSet();
        foreach (var cd in comments)
        {
            if (cd.ParentRef is { } parent && !present.Contains(parent))
                cd.ParentRef = null;
        }
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
        RemoveCommentElems(_root.Elements(PNs + "cm")
            .Where(e => int.Parse(e.Attribute("authorId")?.Value ?? "-1") == authorId));
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
        if (doc.Root is null)
            return null;

        var part = new CommentsPart(doc.Root);
        part.DetachOrphanedReplies();
        return part;
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
    /// reference to a comment the file does not have. <see cref="DetachOrphanedReplies"/> already
    /// does this on every removal; it is repeated here on the copy that is written.
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
