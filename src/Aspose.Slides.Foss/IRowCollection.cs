namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of table rows.
/// </summary>
public interface IRowCollection
{
    /// <summary>
    /// Gets the row at the specified index.
    /// </summary>
    IRow this[int index] { get; }

    /// <summary>
    /// Gets the number of rows.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Gets a shallow copy of all rows as a list.
    /// </summary>
    IList<IRow> AsICollection { get; }

    /// <summary>
    /// Gets an enumerable over the rows.
    /// </summary>
    IEnumerable<IRow> AsIEnumerable { get; }

    /// <summary>
    /// Clones the template row and appends the result to the end of the collection.
    /// When <paramref name="withAttachedRows"/> is <c>true</c>, additional
    /// structurally associated rows (e.g., merged-cell spans) are also cloned and appended.
    /// </summary>
    /// <param name="templ">The template row to clone from.</param>
    /// <param name="withAttachedRows">
    /// When <c>true</c>, also clones attached/associated rows.
    /// </param>
    /// <returns>The list of newly created rows.</returns>
    IList<IRow> AddClone(IRow templ, bool withAttachedRows);

    /// <summary>
    /// Clones the template row and inserts the result at the specified position.
    /// When <paramref name="withAttachedRows"/> is <c>true</c>, additional
    /// structurally associated rows are also cloned and inserted.
    /// </summary>
    /// <param name="index">Zero-based position at which to insert.</param>
    /// <param name="templ">The template row to clone from.</param>
    /// <param name="withAttachedRows">
    /// When <c>true</c>, also clones attached/associated rows.
    /// </param>
    /// <returns>The list of newly created rows.</returns>
    IList<IRow> InsertClone(int index, IRow templ, bool withAttachedRows);

    /// <summary>
    /// Removes the row at the specified index.
    /// When <paramref name="withAttachedRows"/> is <c>true</c>, also removes
    /// dependent/attached rows.
    /// </summary>
    /// <param name="firstRowIndex">Zero-based index of the row to remove.</param>
    /// <param name="withAttachedRows">
    /// When <c>true</c>, also removes attached/associated rows.
    /// </param>
    void RemoveAt(int firstRowIndex, bool withAttachedRows);
}
