namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of table columns.
/// </summary>
public interface IColumnCollection
{
    /// <summary>
    /// Gets the column at the specified index.
    /// </summary>
    IColumn this[int index] { get; }

    /// <summary>
    /// Gets the number of columns.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Gets a shallow copy of all columns as a list.
    /// </summary>
    IList<IColumn> AsICollection { get; }

    /// <summary>
    /// Gets an enumerable over the columns.
    /// </summary>
    IEnumerable<IColumn> AsIEnumerable { get; }

    /// <summary>
    /// Clones the template column and appends the result to the end of the collection.
    /// When <paramref name="withAttachedColumns"/> is <c>true</c>, additional
    /// structurally associated columns are also cloned and appended.
    /// </summary>
    /// <param name="templ">The template column to clone from.</param>
    /// <param name="withAttachedColumns">
    /// When <c>true</c>, also clones attached/associated columns.
    /// </param>
    /// <returns>The list of newly created columns.</returns>
    IList<IColumn> AddClone(IColumn templ, bool withAttachedColumns);

    /// <summary>
    /// Clones the template column and inserts the result at the specified position.
    /// When <paramref name="withAttachedColumns"/> is <c>true</c>, additional
    /// structurally associated columns are also cloned and inserted.
    /// </summary>
    /// <param name="index">Zero-based position at which to insert.</param>
    /// <param name="templ">The template column to clone from.</param>
    /// <param name="withAttachedColumns">
    /// When <c>true</c>, also clones attached/associated columns.
    /// </param>
    /// <returns>The list of newly created columns.</returns>
    IList<IColumn> InsertClone(int index, IColumn templ, bool withAttachedColumns);

    /// <summary>
    /// Removes the column at the specified index.
    /// </summary>
    void RemoveAt(int index, bool withAttachedRows);
}
