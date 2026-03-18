using Aspose.Slides.Foss.Drawing;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of comments.
/// </summary>
public interface ICommentCollection
{
    /// <summary>
    /// Returns this collection as a generic list. Returns a new list each time.
    /// </summary>
    IList<IComment> AsICollection { get; }

    /// <summary>
    /// Returns this collection as an enumerable. Returns a new sequence each time.
    /// </summary>
    IEnumerable<IComment> AsIEnumerable { get; }

    /// <summary>
    /// Returns all comments in this collection as a new list.
    /// </summary>
    /// <returns>An array of all comments.</returns>
    IComment[] ToArray();

    /// <summary>
    /// Returns a range of comments from this collection.
    /// </summary>
    /// <param name="start">The zero-based start index.</param>
    /// <param name="count">The number of comments to return.</param>
    /// <returns>An array of comments in the specified range.</returns>
    IComment[] ToArray(int start, int count);

    /// <summary>
    /// Adds a new comment to the collection.
    /// </summary>
    /// <param name="text">The comment text.</param>
    /// <param name="slide">The target slide.</param>
    /// <param name="position">The position on the slide.</param>
    /// <param name="creationTime">The creation timestamp.</param>
    /// <returns>The newly created comment.</returns>
    IComment AddComment(string text, ISlide slide, PointF position, DateTime creationTime);

    /// <summary>
    /// Inserts a new comment at the specified position in the collection.
    /// </summary>
    /// <param name="index">The zero-based insertion index within the comments part.</param>
    /// <param name="text">The comment text.</param>
    /// <param name="slide">The target slide.</param>
    /// <param name="position">The position on the slide.</param>
    /// <param name="creationTime">The creation timestamp.</param>
    /// <returns>The newly created comment.</returns>
    IComment InsertComment(int index, string text, ISlide slide, PointF position, DateTime creationTime);

    /// <summary>
    /// Removes the comment at the specified positional index in the flat enumeration.
    /// No-op if the index is out of bounds.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Removes the specified comment by delegating to its own <see cref="IComment.Remove"/> method.
    /// </summary>
    /// <param name="comment">The comment to remove.</param>
    void Remove(IComment comment);

    /// <summary>
    /// Removes all comments from the collection.
    /// </summary>
    void Clear();

    /// <summary>
    /// Gets the comment at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the comment to get.</param>
    /// <returns>The comment at the specified index.</returns>
    IComment this[int index] { get; }

    /// <summary>
    /// Gets the number of comments in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Finds a comment by its persistent author-scoped comment index.
    /// </summary>
    /// <param name="idx">The comment index (as stored in XML).</param>
    /// <returns>The matching comment, or <c>null</c> if not found.</returns>
    IComment? FindCommentByIdx(int idx);
}
