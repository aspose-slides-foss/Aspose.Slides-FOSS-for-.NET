namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Internal implementation of <see cref="ICommentAuthor"/> used when constructing parent comment wrappers.
/// </summary>
internal sealed class CommentAuthorImpl : ICommentAuthor
{
    private readonly AuthorData _data;

    internal CommentAuthorImpl(AuthorData data)
    {
        _data = data;
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
    public ICommentCollection Comments =>
        throw new NotSupportedException("Comments are not available on internal author wrappers.");

    /// <inheritdoc/>
    public void Remove() =>
        throw new NotSupportedException("Remove is not available on internal author wrappers.");
}
