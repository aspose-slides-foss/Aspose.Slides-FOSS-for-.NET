using System.Collections;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Internal base class for all collection types.
/// Concrete subclasses must implement <see cref="Count"/> and the indexer.
/// In return they get <see cref="Length"/>, <see cref="GetEnumerator"/>,
/// and <see cref="Contains"/> for free.
/// </summary>
/// <typeparam name="T">The element type of the collection.</typeparam>
public abstract class BaseCollection<T> : IEnumerable<T>
{
    /// <summary>
    /// Gets the number of elements in the collection.
    /// </summary>
    protected abstract int Count { get; }

    /// <summary>
    /// Gets the element at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the element.</param>
    protected abstract T GetItem(int index);

    /// <summary>
    /// Gets the number of elements in the collection. Read-only.
    /// </summary>
    public int Length => Count;

    /// <summary>
    /// Determines whether the collection contains a specific item,
    /// using reference equality first, then value equality.
    /// </summary>
    /// <param name="item">The item to locate.</param>
    /// <returns><c>true</c> if the item is found; otherwise <c>false</c>.</returns>
    public bool Contains(T item)
    {
        for (var i = 0; i < Count; i++)
        {
            var current = GetItem(i);
            if (ReferenceEquals(current, item) || Equals(current, item))
                return true;
        }

        return false;
    }

    /// <inheritdoc />
    public IEnumerator<T> GetEnumerator()
    {
        for (var i = 0; i < Count; i++)
            yield return GetItem(i);
    }

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
