using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a comment on a slide.
/// </summary>
public sealed class Comment : IComment
{
    private CommentData _data = null!;
    private CommentsPart _commentsPart = null!;
    private CommentAuthorsPart _authorsPart = null!;
    private ISlide _slideRef = null!;
    private ICommentAuthor _authorRef = null!;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        CommentData data,
        CommentsPart commentsPart,
        CommentAuthorsPart authorsPart,
        ISlide slide,
        ICommentAuthor author)
    {
        _data = data;
        _commentsPart = commentsPart;
        _authorsPart = authorsPart;
        _slideRef = slide;
        _authorRef = author;
    }

    /// <inheritdoc/>
    public string Text
    {
        get => _data.Text;
        set => _data.Text = value;
    }

    /// <inheritdoc/>
    public DateTime? CreatedTime
    {
        get => DateTimeHelpers.StrToDt(_data.DtStr);
        set => _data.DtStr = value is not null ? DateTimeHelpers.DtToStr(value.Value) : "";
    }

    /// <inheritdoc/>
    public ISlide Slide => _slideRef;

    /// <inheritdoc/>
    public ICommentAuthor Author => _authorRef;

    /// <inheritdoc/>
    public PointF Position
    {
        get => new(_data.PosX, _data.PosY);
        set
        {
            _data.PosX = value.X;
            _data.PosY = value.Y;
        }
    }

    /// <inheritdoc/>
    public IComment? ParentComment
    {
        get
        {
            var parentId = _data.ParentCmId;
            if (parentId is null)
                return null;

            var cd = _commentsPart.FindCommentByIdxAll(parentId.Value);
            if (cd is null)
                return null;

            var authorData = _authorsPart.FindAuthorById(cd.AuthorId);
            if (authorData is null)
                return null;

            var author = new CommentAuthorImpl(authorData);

            var parent = new Comment();
            parent.InitInternal(
                data: cd,
                commentsPart: _commentsPart,
                authorsPart: _authorsPart,
                slide: _slideRef,
                author: author);
            return parent;
        }
        set
        {
            if (value is null)
            {
                _data.ParentCmId = null;
            }
            else
            {
                var other = (Comment)value;
                _data.ParentCmId = other._data.Idx;
            }
        }
    }

    /// <inheritdoc/>
    public void Remove()
    {
        var myIdx = _data.Idx;
        var myAuthorId = _data.AuthorId;

        // Collect direct replies by the same author
        var toRemove = new List<System.Xml.Linq.XElement>();
        foreach (var cd in _commentsPart.GetCommentsByAuthor(myAuthorId))
        {
            if (cd.ParentCmId == myIdx)
                toRemove.Add(cd.Elem);
        }

        // Remove replies
        foreach (var elem in toRemove)
            _commentsPart.RemoveCommentElem(elem);

        // Remove this comment
        _commentsPart.RemoveCommentElem(_data.Elem);

        // Persist changes
        _commentsPart.Save();
    }
}
