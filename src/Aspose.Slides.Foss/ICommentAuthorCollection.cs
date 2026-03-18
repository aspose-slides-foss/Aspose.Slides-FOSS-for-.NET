using System.Collections.Generic;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of comment authors.
/// </summary>
public interface ICommentAuthorCollection
{
    /// <summary>
    /// Returns the collection as a plain list. Read-only.
    /// </summary>
    IList<ICommentAuthor> AsICollection { get; }

    /// <summary>
    /// Returns the collection as an enumerable. Read-only.
    /// </summary>
    IEnumerable<ICommentAuthor> AsIEnumerable { get; }

    /// <summary>
    /// Adds a new author to the collection.
    /// </summary>
    /// <param name="name">The name of the author.</param>
    /// <param name="initials">The initials of the author.</param>
    /// <returns>The newly created comment author.</returns>
    ICommentAuthor AddAuthor(string name, string initials);

    /// <summary>
    /// Returns all authors as an array.
    /// </summary>
    /// <returns>An array of all comment authors.</returns>
    ICommentAuthor[] ToArray();

    /// <summary>
    /// Finds authors by name (exact, case-sensitive match).
    /// </summary>
    /// <param name="name">The name to search for.</param>
    /// <returns>An array of matching authors.</returns>
    ICommentAuthor[] FindByName(string name);

    /// <summary>
    /// Finds authors by name and initials (exact, case-sensitive match).
    /// </summary>
    /// <param name="name">The name to search for.</param>
    /// <param name="initials">The initials to search for.</param>
    /// <returns>An array of matching authors.</returns>
    ICommentAuthor[] FindByNameAndInitials(string name, string initials);

    /// <summary>
    /// Removes the author at the specified index.
    /// If the index is out of bounds, the method silently does nothing.
    /// All comments by the author are removed before the author record is deleted.
    /// </summary>
    /// <param name="index">The zero-based index of the author to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Removes the specified author from the collection.
    /// </summary>
    /// <param name="author">The author to remove.</param>
    void Remove(ICommentAuthor author);

    /// <summary>
    /// Removes all authors and their associated comments.
    /// </summary>
    void Clear();

    /// <summary>
    /// Gets the author at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The comment author at the specified index.</returns>
    /// <exception cref="IndexOutOfRangeException">Thrown when the index is out of range.</exception>
    ICommentAuthor this[int index] { get; }

    /// <summary>
    /// Gets the number of authors in the collection.
    /// </summary>
    int Count { get; }
}
