using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of columns in a table. Maintains synchronization between
/// the <c>tblGrid</c> column definitions and per-row cell elements in the XML tree.
/// </summary>
public sealed class ColumnCollection : IColumnCollection, IEnumerable<IColumn>
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private XElement? _tblElement;
    private XElement? _tblGridElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;
    private ITable? _table;
    private List<Column> _columns = new();

    /// <summary>
    /// Initializes internal state from the provided table XML elements.
    /// </summary>
    internal void InitInternal(XElement tblElement, XElement tblGridElement, SlidePart? slidePart, IBaseSlide? parentSlide, ITable? table)
    {
        _tblElement = tblElement;
        _tblGridElement = tblGridElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _table = table;
        Rebuild();
    }

    /// <inheritdoc/>
    public IColumn this[int index] => _columns[index];

    /// <inheritdoc/>
    public int Count => _columns.Count;

    /// <summary>
    /// Gets a shallow copy of all columns as a list.
    /// </summary>
    public IList<IColumn> AsICollection => new List<IColumn>(_columns);

    /// <summary>
    /// Gets an enumerable over the columns.
    /// </summary>
    public IEnumerable<IColumn> AsIEnumerable => _columns;

    /// <inheritdoc/>
    public IList<IColumn> AddClone(IColumn templ, bool withAttachedColumns)
    {
        if (_tblElement is null || _tblGridElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var srcCol = (Column)templ;

        // Deep-copy the template's gridCol element and append to tblGrid.
        var newGridCol = new XElement(srcCol.GridColElement!);
        _tblGridElement.Add(newGridCol);

        // For each row, deep-copy the tc at the template column's index, or create empty tc.
        var templateColIndex = srcCol.ColIndex;
        foreach (var tr in _tblElement.Elements(ANs + "tr"))
        {
            var tcs = tr.Elements(ANs + "tc").ToList();
            var newTc = templateColIndex < tcs.Count
                ? new XElement(tcs[templateColIndex])
                : MakeEmptyTc();
            tr.Add(newTc);
        }

        var insertedAt = _columns.Count;
        Rebuild();
        _slidePart?.Save();

        // Collect all newly appended columns (one unless withAttachedColumns added more).
        var result = new List<IColumn>();
        for (var i = insertedAt; i < _columns.Count; i++)
            result.Add(_columns[i]);
        return result;
    }

    /// <inheritdoc/>
    public IList<IColumn> InsertClone(int index, IColumn templ, bool withAttachedColumns)
    {
        if (_tblElement is null || _tblGridElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var previousCount = _columns.Count;
        var srcCol = (Column)templ;

        // Deep-copy the template's gridCol element.
        var newGridCol = new XElement(srcCol.GridColElement!);
        var gridCols = _tblGridElement.Elements(ANs + "gridCol").ToList();
        if (index < gridCols.Count)
            gridCols[index].AddBeforeSelf(newGridCol);
        else
            _tblGridElement.Add(newGridCol);

        // For each row, deep-copy the tc at the template column's index, or create empty tc.
        var templateColIndex = srcCol.ColIndex;
        foreach (var tr in _tblElement.Elements(ANs + "tr"))
        {
            var tcs = tr.Elements(ANs + "tc").ToList();
            var newTc = templateColIndex < tcs.Count
                ? new XElement(tcs[templateColIndex])
                : MakeEmptyTc();

            if (index < tcs.Count)
                tcs[index].AddBeforeSelf(newTc);
            else
                tr.Add(newTc);
        }

        Rebuild();
        _slidePart?.Save();

        // Collect all newly inserted columns starting at the insertion point.
        var insertedCount = _columns.Count - previousCount;
        var result = new List<IColumn>(insertedCount);
        for (var i = index; i < index + insertedCount; i++)
            result.Add(_columns[i]);
        return result;
    }

    /// <inheritdoc/>
    public void RemoveAt(int index, bool withAttachedRows)
    {
        if (_tblElement is null || _tblGridElement is null)
            throw new InvalidOperationException("Table element not initialized.");

        var gridCols = _tblGridElement.Elements(ANs + "gridCol").ToList();
        if (index < 0 || index >= gridCols.Count)
            throw new IndexOutOfRangeException($"Column index {index} out of range");

        gridCols[index].Remove();

        foreach (var tr in _tblElement.Elements(ANs + "tr"))
        {
            var tcs = tr.Elements(ANs + "tc").ToList();
            if (index < tcs.Count)
                tcs[index].Remove();
        }

        Rebuild();
        _slidePart?.Save();
    }

    /// <inheritdoc/>
    public IEnumerator<IColumn> GetEnumerator() => _columns.Cast<IColumn>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private void Rebuild()
    {
        _columns.Clear();
        if (_tblElement is null || _tblGridElement is null) return;

        var gridCols = _tblGridElement.Elements(ANs + "gridCol").ToList();
        var rows = _tblElement.Elements(ANs + "tr").ToList();

        for (var colIdx = 0; colIdx < gridCols.Count; colIdx++)
        {
            var cells = new List<Cell>();
            for (var rowIdx = 0; rowIdx < rows.Count; rowIdx++)
            {
                var tcs = rows[rowIdx].Elements(ANs + "tc").ToList();
                if (colIdx < tcs.Count)
                {
                    var cell = new Cell();
                    cell.InitInternal(tcs[colIdx], rowIdx, colIdx, _slidePart, _parentSlide, _table!);
                    cells.Add(cell);
                }
            }

            var col = new Column();
            col.InitInternal(gridCols[colIdx], colIdx, cells, _slidePart, _parentSlide);
            _columns.Add(col);
        }
    }

    private static XElement MakeEmptyTc()
    {
        return new XElement(ANs + "tc",
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "endParaRPr"))),
            new XElement(ANs + "tcPr"));
    }
}
