namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of portions.
/// </summary>
public interface IPortionCollection
{
    /// <summary>
    /// Gets the portion at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the portion to get.</param>
    IPortion this[int index] { get; }

    /// <summary>
    /// Gets the number of elements actually contained in the collection. Read-only.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Returns <see cref="IEnumerable{IPortion}"/> interface. Read-only.
    /// </summary>
    IEnumerable<IPortion> AsIEnumerable { get; }

    /// <summary>
    /// Adds a portion to the end of the collection.
    /// </summary>
    /// <param name="value">The portion to add.</param>
    void Add(IPortion value);

    /// <summary>
    /// Determines the index of a specific portion in the collection.
    /// </summary>
    /// <param name="item">The portion to locate.</param>
    /// <returns>The zero-based index of the portion if found; otherwise, -1.</returns>
    int IndexOf(IPortion item);

    /// <summary>
    /// Inserts a portion into the collection at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index at which the portion should be inserted.</param>
    /// <param name="value">The portion to insert.</param>
    void Insert(int index, IPortion value);

    /// <summary>
    /// Removes all portions from the collection.
    /// </summary>
    void Clear();

    /// <summary>
    /// Determines whether the collection contains a specific portion.
    /// </summary>
    /// <param name="item">The portion to locate.</param>
    /// <returns><c>true</c> if the portion is found; otherwise, <c>false</c>.</returns>
    bool Contains(IPortion item);

    /// <summary>
    /// Removes the first occurrence of a specific portion from the collection.
    /// </summary>
    /// <param name="item">The portion to remove.</param>
    /// <returns><c>true</c> if the portion was found and removed; otherwise, <c>false</c>.</returns>
    bool Remove(IPortion item);

    /// <summary>
    /// Removes the portion at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the portion to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Gets a value indicating whether the collection is read-only. Read-only.
    /// </summary>
    bool IsReadOnly { get; }
}
