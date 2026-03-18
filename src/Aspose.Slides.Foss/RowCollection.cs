using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of rows in a table. Maintains synchronization between
/// the <c>tr</c> row elements in the XML tree.
/// </summary>
public sealed class RowCollection : IRowCollection, IEnumerable<IRow>
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private XElement? _tblElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;
    private ITable? _table;
    private List<Row> _rows = new();

    /// <summary>
    /// Initializes internal state from the provided table XML elements.
    /// </summary>
    internal void InitInternal(XElement tblElement, SlidePart? slidePart, IBaseSlide? parentSlide, ITable? table)
    {
        _tblElement = tblElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _table = table;
        Rebuild();
    }

    /// <inheritdoc/>
    public IRow this[int index] => _rows[index];

    /// <inheritdoc/>
    public int Count => _rows.Count;

    /// <summary>
    /// Gets a shallow copy of all rows as a list.
    /// </summary>
    public IList<IRow> AsICollection => new List<IRow>(_rows);

    /// <summary>
    /// Gets an enumerable over the rows.
    /// </summary>
    public IEnumerable<IRow> AsIEnumerable => _rows;

    /// <inheritdoc/>
    public IList<IRow> AddClone(IRow templ, bool withAttachedRows)
    {
        if (_tblElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var srcRow = (Row)templ;

        // Deep-copy the template's tr element and append to the table.
        var newTr = new XElement(srcRow.TrElement!);
        _tblElement.Add(newTr);

        var insertedAt = _rows.Count;
        Rebuild();
        _slidePart?.Save();

        // Collect all newly appended rows.
        var result = new List<IRow>();
        for (var i = insertedAt; i < _rows.Count; i++)
            result.Add(_rows[i]);
        return result;
    }

    /// <inheritdoc/>
    public IList<IRow> InsertClone(int index, IRow templ, bool withAttachedRows)
    {
        if (_tblElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var previousCount = _rows.Count;
        var srcRow = (Row)templ;

        // Deep-copy the template's tr element.
        var newTr = new XElement(srcRow.TrElement!);
        var trs = _tblElement.Elements(ANs + "tr").ToList();
        if (index < trs.Count)
            trs[index].AddBeforeSelf(newTr);
        else
            _tblElement.Add(newTr);

        Rebuild();
        _slidePart?.Save();

        // Collect all newly inserted rows starting at the insertion point.
        var insertedCount = _rows.Count - previousCount;
        var result = new List<IRow>(insertedCount);
        for (var i = index; i < index + insertedCount; i++)
            result.Add(_rows[i]);
        return result;
    }

    /// <inheritdoc/>
    public void RemoveAt(int firstRowIndex, bool withAttachedRows)
    {
        if (_tblElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var trs = _tblElement.Elements(ANs + "tr").ToList();
        if (firstRowIndex < 0 || firstRowIndex >= trs.Count)
            throw new IndexOutOfRangeException($"Row index {firstRowIndex} out of range");

        trs[firstRowIndex].Remove();

        Rebuild();
        _slidePart?.Save();
    }

    /// <inheritdoc/>
    public IEnumerator<IRow> GetEnumerator() => _rows.Cast<IRow>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private void Rebuild()
    {
        _rows.Clear();
        if (_tblElement is null) return;

        var trs = _tblElement.Elements(ANs + "tr").ToList();

        for (var rowIdx = 0; rowIdx < trs.Count; rowIdx++)
        {
            var tr = trs[rowIdx];
            var cells = new List<Cell>();
            var tcElements = tr.Elements(ANs + "tc").ToList();
            for (var colIdx = 0; colIdx < tcElements.Count; colIdx++)
            {
                var cell = new Cell();
                cell.InitInternal(tcElements[colIdx], rowIdx, colIdx, _slidePart, _parentSlide, _table!);
                cells.Add(cell);
            }

            var row = new Row();
            row.InitInternal(tr, rowIdx, cells, _slidePart, _parentSlide);
            _rows.Add(row);
        }
    }
}
