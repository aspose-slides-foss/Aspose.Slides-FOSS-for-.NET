using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of comment authors backed by a <see cref="CommentAuthorsPart"/>.
/// Every access rebuilds wrapper objects from the underlying part data; no caching is performed.
/// </summary>
public sealed class CommentAuthorCollection : ICommentAuthorCollection
{
    private CommentAuthorsPart _authorsPart = null!;
    private OpcPackage _package = null!;
    private Presentation? _presentation;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        CommentAuthorsPart authorsPart,
        OpcPackage package,
        Presentation? presentation = null)
    {
        _authorsPart = authorsPart;
        _package = package;
        _presentation = presentation;
    }

    /// <inheritdoc/>
    public IList<ICommentAuthor> AsICollection => ToArray();

    /// <inheritdoc/>
    public IEnumerable<ICommentAuthor> AsIEnumerable => ToArray();

    /// <inheritdoc/>
    public ICommentAuthor AddAuthor(string name, string initials)
    {
        var data = _authorsPart.AddAuthor(name, initials);
        return BuildAuthor(data);
    }

    /// <inheritdoc/>
    public ICommentAuthor[] ToArray()
    {
        return _authorsPart.GetAuthors()
            .Select(BuildAuthor)
            .ToArray();
    }

    /// <inheritdoc/>
    public ICommentAuthor[] FindByName(string name)
    {
        return _authorsPart.GetAuthors()
            .Where(d => d.Name == name)
            .Select(BuildAuthor)
            .ToArray();
    }

    /// <inheritdoc/>
    public ICommentAuthor[] FindByNameAndInitials(string name, string initials)
    {
        return _authorsPart.GetAuthors()
            .Where(d => d.Name == name && d.Initials == initials)
            .Select(BuildAuthor)
            .ToArray();
    }

    /// <inheritdoc/>
    public void RemoveAt(int index)
    {
        var authors = _authorsPart.GetAuthors();
        if (index < 0 || index >= authors.Count)
            return;

        var data = authors[index];
        var author = BuildAuthor(data);
        author.Comments.Clear();
        _authorsPart.RemoveAuthor(data.Id);
    }

    /// <inheritdoc/>
    public void Remove(ICommentAuthor author)
    {
        author.Remove();
    }

    /// <inheritdoc/>
    public void Clear()
    {
        foreach (var data in _authorsPart.GetAuthors())
        {
            var author = BuildAuthor(data);
            author.Comments.Clear();
        }

        _authorsPart.Clear();
    }

    /// <summary>
    /// Gets the author at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The comment author at the specified index.</returns>
    /// <exception cref="IndexOutOfRangeException">Thrown when the index is out of range.</exception>
    public ICommentAuthor this[int index]
    {
        get
        {
            var authors = _authorsPart.GetAuthors();
            if (index < 0 || index >= authors.Count)
                throw new IndexOutOfRangeException($"Index {index} out of range");

            return BuildAuthor(authors[index]);
        }
    }

    /// <summary>
    /// Gets the number of authors in the collection.
    /// </summary>
    public int Count => _authorsPart.GetAuthors().Count;

    private ICommentAuthor BuildAuthor(AuthorData data)
    {
        var author = new CommentAuthor();
        author.InitInternal(data, _authorsPart, _package, _presentation);
        return author;
    }
}
