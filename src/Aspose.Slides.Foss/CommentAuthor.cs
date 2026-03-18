using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an author of comments in a presentation.
/// </summary>
public sealed class CommentAuthor : ICommentAuthor
{
    private AuthorData _data = null!;
    private CommentAuthorsPart _authorsPart = null!;
    private OpcPackage _package = null!;
    private Presentation? _presentation;
    private CommentCollection? _commentsCache;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(
        AuthorData data,
        CommentAuthorsPart authorsPart,
        OpcPackage package,
        Presentation? presentation = null)
    {
        _data = data;
        _authorsPart = authorsPart;
        _package = package;
        _presentation = presentation;
        _commentsCache = null;
    }

    /// <inheritdoc/>
    public string Name
    {
        get => _data.Name;
        set => _data.Name = value;
    }

    /// <inheritdoc/>
    public string Initials
    {
        get => _data.Initials;
        set => _data.Initials = value;
    }

    /// <inheritdoc/>
    public ICommentCollection Comments
    {
        get
        {
            if (_commentsCache is null)
            {
                var collection = new CommentCollection();
                collection.InitInternal(_data, _authorsPart, _package, _presentation);
                _commentsCache = collection;
            }

            return _commentsCache;
        }
    }

    /// <inheritdoc/>
    public void Remove()
    {
        Comments.Clear();
        _authorsPart.RemoveAuthor(_data.Id);
    }
}
