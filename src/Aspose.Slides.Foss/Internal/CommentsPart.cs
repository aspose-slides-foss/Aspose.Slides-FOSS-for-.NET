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

    internal CommentData? FindCommentByIdxAll(int idx)
    {
        var elem = _root.Elements(PNs + "cm")
            .FirstOrDefault(e => int.Parse(e.Attribute("idx")?.Value ?? "-1") == idx);
        return elem is not null ? new CommentData(elem) : null;
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
    /// The attribute this class uses in memory to remember which comment a reply answers.
    /// </summary>
    /// <remarks>
    /// CT_Comment declares <c>authorId</c>, <c>dt</c> and <c>idx</c> and nothing else, so this
    /// attribute makes the part schema-invalid and no consumer renders a thread from it. Threading
    /// in PowerPoint is a separate <c>ppt/threadedComments/</c> part, which this library does not
    /// write. The marker is therefore stripped on the way out: <see cref="Comment.ParentComment"/>
    /// answers from memory, and the file carries flat comments, which is what it really has.
    /// </remarks>
    internal const string ParentMarkerAttribute = "parentCmId";

    /// <summary>
    /// Flushes the comments XML back to the OPC package at the given part name.
    /// </summary>
    internal void Flush(OpcPackage package, string partName)
    {
        var serialized = new XElement(_root);
        foreach (var marker in serialized.DescendantsAndSelf()
                     .Attributes(ParentMarkerAttribute).ToList())
        {
            marker.Remove();
        }

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), serialized);
        using var ms = new MemoryStream();
        doc.Save(ms);
        package.SetPart(partName, ms.ToArray());
    }

    private static XElement CreateCommentElement(int authorId, int idx, string text, float posX, float posY, string dtStr)
    {
        const int cmToEmu = 360000;
        return new XElement(PNs + "cm",
            new XAttribute("authorId", authorId),
            new XAttribute("idx", idx),
            new XAttribute("dt", dtStr),
            new XElement(PNs + "pos",
                new XAttribute("x", (int)Math.Round(posX * cmToEmu)),
                new XAttribute("y", (int)Math.Round(posY * cmToEmu))),
            new XElement(PNs + "text", text));
    }
}
