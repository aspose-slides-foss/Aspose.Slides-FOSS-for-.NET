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
            var parentRef = _data.ParentRef;
            if (parentRef is null)
                return null;

            var cd = _commentsPart.FindComment(parentRef.Value.AuthorId, parentRef.Value.Idx);
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
                _data.ParentRef = null;
                return;
            }

            if (value is not Comment other)
                throw new ArgumentException("The parent must be a comment of this presentation.", nameof(value));
            if (ReferenceEquals(other._data.Elem, _data.Elem))
                throw new ArgumentException("A comment cannot reply to itself.", nameof(value));
            // PowerPoint names a parent within the slide's own comment list, so a reply and the
            // comment it answers must be on the same slide.
            if (!ReferenceEquals(other._commentsPart, _commentsPart))
                throw new ArgumentException("A reply must be on the same slide as the comment it answers.", nameof(value));

            other._data.EnsureThreadingInfo();
            _data.ParentRef = (other._data.AuthorId, other._data.Idx);
        }
    }

    /// <inheritdoc/>
    public void Remove()
    {
        // Direct replies, by any author: each names this comment by its author and index.
        foreach (var reply in _commentsPart.GetReplies(_data))
            _commentsPart.RemoveCommentElem(reply.Elem);

        // Remove this comment
        _commentsPart.RemoveCommentElem(_data.Elem);

        // Persist changes
        _commentsPart.Save();
    }
}
