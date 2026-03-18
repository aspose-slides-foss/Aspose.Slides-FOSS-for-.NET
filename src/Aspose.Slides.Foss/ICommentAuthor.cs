namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an author of comments.
/// </summary>
public interface ICommentAuthor
{
    /// <summary>
    /// Returns or sets the author's name.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Returns or sets the author's initials.
    /// </summary>
    string Initials { get; set; }

    /// <summary>
    /// Returns the collection of comments made by this author. Read-only.
    /// </summary>
    ICommentCollection Comments { get; }

    /// <summary>
    /// Removes this author and all associated comments.
    /// </summary>
    void Remove();
}
