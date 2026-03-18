namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of paragraphs.
/// </summary>
public interface IParagraphCollection
{
    /// <summary>
    /// Gets the paragraph at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the paragraph to get.</param>
    IParagraph this[int index] { get; }

    /// <summary>
    /// Gets the number of elements actually contained in the collection. Read-only.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Allows to get base <see cref="ISlideComponent"/> interface. Read-only.
    /// </summary>
    ISlideComponent AsISlideComponent { get; }

    /// <summary>
    /// Returns <see cref="IEnumerable{IParagraph}"/> interface. Read-only.
    /// </summary>
    IEnumerable<IParagraph> AsIEnumerable { get; }

    /// <summary>
    /// Adds a paragraph to the end of the collection.
    /// </summary>
    /// <param name="value">The paragraph to add.</param>
    void Add(IParagraph value);

    /// <summary>
    /// Inserts a paragraph into the collection at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index at which the paragraph should be inserted.</param>
    /// <param name="value">The paragraph to insert.</param>
    void Insert(int index, IParagraph value);

    /// <summary>
    /// Removes all paragraphs from the collection.
    /// </summary>
    void Clear();

    /// <summary>
    /// Removes the paragraph at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the paragraph to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Removes the first occurrence of a specific paragraph from the collection.
    /// </summary>
    /// <param name="item">The paragraph to remove.</param>
    /// <returns><c>true</c> if the paragraph was found and removed; otherwise, <c>false</c>.</returns>
    bool Remove(IParagraph item);

    /// <summary>
    /// Determines whether the collection contains a specific paragraph.
    /// </summary>
    /// <param name="item">The paragraph to locate.</param>
    /// <returns><c>true</c> if the paragraph is found; otherwise, <c>false</c>.</returns>
    bool Contains(IParagraph item);

    /// <summary>
    /// Determines the index of a specific paragraph in the collection.
    /// </summary>
    /// <param name="item">The paragraph to locate.</param>
    /// <returns>The zero-based index of the paragraph if found; otherwise, -1.</returns>
    int IndexOf(IParagraph item);

    /// <summary>
    /// Gets a value indicating whether the collection is read-only. Read-only.
    /// </summary>
    bool IsReadOnly { get; }
}
