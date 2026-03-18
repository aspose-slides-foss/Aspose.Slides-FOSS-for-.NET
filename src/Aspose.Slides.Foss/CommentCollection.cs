using System.Text.RegularExpressions;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of comments authored by a single author across all slides in a presentation.
/// Every read operation rebuilds the comment list from the package; no list-level caching is performed.
/// </summary>
public sealed partial class CommentCollection : ICommentCollection
{
    private static readonly Regex SlidePartPattern = new(@"^ppt/slides/slide(\d+)\.xml$", RegexOptions.IgnoreCase);

    private AuthorData _data = null!;
    private CommentAuthorsPart _authorsPart = null!;
    private OpcPackage _package = null!;
    private Presentation? _presentation;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        AuthorData data,
        CommentAuthorsPart authorsPart,
        OpcPackage package,
        Presentation? presentation)
    {
        _data = data;
        _authorsPart = authorsPart;
        _package = package;
        _presentation = presentation;
    }

    /// <summary>
    /// Gets the comment at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the comment.</param>
    /// <returns>The comment at the specified index.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="index"/> is out of range.</exception>
    public IComment this[int index]
    {
        get
        {
            var allEntries = CollectMyComments();
            if (index < 0 || index >= allEntries.Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index {index} out of range");
            var (cd, cp, pn) = allEntries[index];
            return BuildComment(cd, cp, ResolveSlide(pn));
        }
    }

    /// <summary>
    /// Gets the number of comments in this collection.
    /// </summary>
    public int Count => CollectMyComments().Count;

    /// <inheritdoc/>
    public IList<IComment> AsICollection => ToArray();

    /// <inheritdoc/>
    public IEnumerable<IComment> AsIEnumerable => ToArray();

    /// <inheritdoc/>
    public IComment[] ToArray()
    {
        var collected = CollectMyComments();
        var result = new List<IComment>(collected.Count);
        foreach (var (cd, cp, slidePartName) in collected)
        {
            var slide = ResolveSlide(slidePartName);
            result.Add(BuildComment(cd, cp, slide));
        }
        return [.. result];
    }

    /// <inheritdoc/>
    public IComment[] ToArray(int start, int count)
    {
        var all = ToArray();
        var end = Math.Min(start + count, all.Length);
        if (start >= all.Length || start < 0)
            return [];
        return all[start..end];
    }

    /// <inheritdoc/>
    public IComment AddComment(string text, ISlide slide, PointF position, DateTime creationTime)
    {
        var authorId = _data.Id;
        var idx = _authorsPart.NextCommentIdx(authorId);
        var slidePartName = ((Slide)slide).PartName;
        var cp = GetOrCreateCommentsPart(slidePartName)!;
        var dtStr = DateTimeHelpers.DtToStr(creationTime);

        var cd = cp.AddComment(authorId, idx, text, position.X, position.Y, dtStr);

        cp.Save();
        _authorsPart.SaveAuthors();

        return BuildComment(cd, cp, slide);
    }

    /// <inheritdoc/>
    public IComment InsertComment(int index, string text, ISlide slide, PointF position, DateTime creationTime)
    {
        var authorId = _data.Id;
        var idx = _authorsPart.NextCommentIdx(authorId);
        var slidePartName = ((Slide)slide).PartName;
        var cp = GetOrCreateCommentsPart(slidePartName)!;
        var dtStr = DateTimeHelpers.DtToStr(creationTime);

        var cd = cp.InsertComment(index, authorId, idx, text, position.X, position.Y, dtStr);

        cp.Save();
        _authorsPart.SaveAuthors();

        return BuildComment(cd, cp, slide);
    }

    /// <inheritdoc/>
    public void RemoveAt(int index)
    {
        var collected = CollectMyComments();
        if (index < 0 || index >= collected.Count)
            return;

        var (cd, cp, _) = collected[index];
        cp.RemoveCommentElem(cd.Elem);
        cp.Save();
    }

    /// <inheritdoc/>
    public void Remove(IComment comment)
    {
        comment.Remove();
    }

    /// <inheritdoc/>
    public void Clear()
    {
        var authorId = _data.Id;
        foreach (var partName in _package.GetPartNames())
        {
            if (!SlidePartPattern.IsMatch(partName))
                continue;

            var commentsPartName = DeriveCommentsPartName(partName);
            var cp = GetOrCreateCommentsPart(commentsPartName, fromSlide: false);
            if (cp is null)
                continue;

            cp.RemoveCommentsByAuthor(authorId);
            cp.Save();
        }
    }

    /// <inheritdoc/>
    public IComment? FindCommentByIdx(int idx)
    {
        foreach (var (cd, cp, slidePartName) in CollectMyComments())
        {
            if (cd.Idx == idx)
            {
                var slide = ResolveSlide(slidePartName);
                return BuildComment(cd, cp, slide);
            }
        }
        return null;
    }

    /// <summary>
    /// Returns all (CommentsPart, slide part name) pairs across slides that have comments by this author.
    /// </summary>
    internal List<(CommentsPart Cp, string SlidePartName)> GetAllCommentsParts()
    {
        var authorId = _data.Id;
        var result = new List<(CommentsPart, string)>();

        foreach (var partName in _package.GetPartNames())
        {
            if (!SlidePartPattern.IsMatch(partName))
                continue;

            var commentsPartName = DeriveCommentsPartName(partName);
            var cp = GetOrCreateCommentsPart(commentsPartName, fromSlide: false);
            if (cp is null)
                continue;

            var comments = cp.GetCommentsByAuthor(authorId);
            if (comments.Count > 0)
                result.Add((cp, partName));
        }

        return result;
    }

    internal List<(CommentData Cd, CommentsPart Cp, string SlidePartName)> CollectMyComments()
    {
        var authorId = _data.Id;
        var result = new List<(CommentData, CommentsPart, string)>();

        foreach (var partName in _package.GetPartNames())
        {
            if (!SlidePartPattern.IsMatch(partName))
                continue;

            var commentsPartName = DeriveCommentsPartName(partName);
            var cp = GetOrCreateCommentsPart(commentsPartName, fromSlide: false);
            if (cp is null)
                continue;

            foreach (var cd in cp.GetCommentsByAuthor(authorId))
            {
                result.Add((cd, cp, partName));
            }
        }

        return result;
    }

    internal CommentsPart? GetOrCreateCommentsPart(string slidePartName, bool fromSlide = true)
    {
        var commentsPartName = fromSlide ? DeriveCommentsPartName(slidePartName) : slidePartName;

        if (_authorsPart.CpCache.TryGetValue(commentsPartName, out var cached))
            return cached;

        var cp = CommentsPart.LoadFromPackage(_package, commentsPartName);
        if (cp is not null)
        {
            _authorsPart.CpCache[commentsPartName] = cp;
            return cp;
        }

        if (!fromSlide)
            return null;

        cp = CommentsPart.CreateEmpty();
        _authorsPart.CpCache[commentsPartName] = cp;
        return cp;
    }

    internal ISlide? ResolveSlide(string partName)
    {
        if (_presentation is null)
            return null;

        foreach (var slide in _presentation.SlidesInternal)
        {
            if (slide is Slide s && string.Equals(s.PartName, partName, StringComparison.OrdinalIgnoreCase))
                return slide;
        }

        return null;
    }

    internal ICommentAuthor GetAuthorObj()
    {
        var author = new CommentAuthor();
        author.InitInternal(_data, _authorsPart, _package, _presentation);
        return author;
    }

    internal Comment BuildComment(CommentData data, CommentsPart cp, ISlide? slide)
    {
        var comment = new Comment();
        comment.InitInternal(
            data: data,
            commentsPart: cp,
            authorsPart: _authorsPart,
            slide: slide!,
            author: GetAuthorObj());
        return comment;
    }

    private static string DeriveCommentsPartName(string slidePartName)
    {
        var match = SlidePartPattern.Match(slidePartName);
        if (match.Success)
            return $"ppt/comments/comment{match.Groups[1].Value}.xml";
        return slidePartName.Replace("ppt/slides/slide", "ppt/comments/comment", StringComparison.Ordinal);
    }
}
