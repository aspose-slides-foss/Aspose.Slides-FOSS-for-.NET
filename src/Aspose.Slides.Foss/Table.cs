using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a table shape on a slide.
/// </summary>
public sealed class Table : GraphicalObject, ITable
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private XElement? _tbl;
    private XElement? _tblPr;
    private XElement? _tblGrid;

    /// <summary>
    /// Initializes internal state from the graphicFrame XML element.
    /// Locates the embedded <c>a:tbl</c>, <c>a:tblPr</c>, and <c>a:tblGrid</c> elements.
    /// </summary>
    internal override void InitInternal(XElement? element, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        base.InitInternal(element, slidePart, parentSlide);

        _tbl = element?.Descendants(ANs + "tbl").FirstOrDefault();
        if (_tbl is not null)
        {
            _tblPr = _tbl.Element(ANs + "tblPr");
            _tblGrid = _tbl.Element(ANs + "tblGrid");
        }
    }

    /// <inheritdoc />
    public override IGraphicalObjectLock? GraphicalObjectLock => null;

    /// <inheritdoc />
    public IRowCollection Rows
    {
        get
        {
            if (_tbl is null)
                return new RowCollection();
            var rc = new RowCollection();
            rc.InitInternal(_tbl, _slidePart, _parentSlide, this);
            return rc;
        }
    }

    /// <inheritdoc />
    public IColumnCollection Columns
    {
        get
        {
            if (_tbl is null || _tblGrid is null)
                return new ColumnCollection();
            var cc = new ColumnCollection();
            cc.InitInternal(_tbl, _tblGrid, _slidePart, _parentSlide, this);
            return cc;
        }
    }

    /// <inheritdoc />
    public ITableFormat TableFormat
    {
        get
        {
            var tf = new TableFormat();
            tf.InitInternal(_tblPr, _slidePart, _parentSlide);
            return tf;
        }
    }

    /// <inheritdoc />
    public TableStylePreset StylePreset
    {
        get
        {
            if (_tblPr is null)
                return TableStylePreset.None;

            var styleEl = FindStyleElement(_tblPr);
            if (styleEl is null)
                return TableStylePreset.None;

            var guid = ReadStyleGuid(styleEl);
            return TableStyleMapping.FromGuid(guid);
        }
        set
        {
            var tblPr = EnsureTblPr();
            var styleEl = FindStyleElement(tblPr);

            if (value == TableStylePreset.None)
            {
                styleEl?.Remove();
            }
            else
            {
                var guid = TableStyleMapping.ToGuid(value);
                if (guid is not null)
                {
                    styleEl?.Remove();
                    var newEl = new XElement(ANs + "tblStyleId") { Value = guid };
                    tblPr.Add(newEl);
                }
            }

            _slidePart?.Save();
        }
    }

    /// <inheritdoc />
    public bool RightToLeft
    {
        get => GetBoolAttr("rtl");
        set => SetBoolAttr("rtl", value);
    }

    /// <inheritdoc />
    public bool FirstRow
    {
        get => GetBoolAttr("firstRow");
        set => SetBoolAttr("firstRow", value);
    }

    /// <inheritdoc />
    public bool FirstCol
    {
        get => GetBoolAttr("firstCol");
        set => SetBoolAttr("firstCol", value);
    }

    /// <inheritdoc />
    public bool LastRow
    {
        get => GetBoolAttr("lastRow");
        set => SetBoolAttr("lastRow", value);
    }

    /// <inheritdoc />
    public bool LastCol
    {
        get => GetBoolAttr("lastCol");
        set => SetBoolAttr("lastCol", value);
    }

    /// <inheritdoc />
    public bool HorizontalBanding
    {
        get => GetBoolAttr("bandRow");
        set => SetBoolAttr("bandRow", value);
    }

    /// <inheritdoc />
    public bool VerticalBanding
    {
        get => GetBoolAttr("bandCol");
        set => SetBoolAttr("bandCol", value);
    }

    /// <inheritdoc />
    public IGraphicalObject AsIGraphicalObject => this;

    /// <inheritdoc />
    public IBulkTextFormattable AsIBulkTextFormattable => this;

    /// <inheritdoc />
    public ICell MergeCells(ICell cell1, ICell cell2, bool allowSplitting)
    {
        if (_tbl is null)
            throw new InvalidOperationException("Table element not initialized.");

        var c1 = (Cell)cell1;
        var c2 = (Cell)cell2;

        var minRow = Math.Min(c1.FirstRowIndex, c2.FirstRowIndex);
        var maxRow = Math.Max(c1.FirstRowIndex, c2.FirstRowIndex);
        var minCol = Math.Min(c1.FirstColumnIndex, c2.FirstColumnIndex);
        var maxCol = Math.Max(c1.FirstColumnIndex, c2.FirstColumnIndex);

        var rowSpan = maxRow - minRow + 1;
        var colSpan = maxCol - minCol + 1;

        var trs = _tbl.Elements(ANs + "tr").ToList();

        for (var r = minRow; r <= maxRow && r < trs.Count; r++)
        {
            var tcs = trs[r].Elements(ANs + "tc").ToList();
            for (var c = minCol; c <= maxCol && c < tcs.Count; c++)
            {
                var tc = tcs[c];
                if (r == minRow && c == minCol)
                {
                    // Origin cell: set spans
                    if (colSpan > 1)
                        tc.SetAttributeValue("gridSpan", colSpan.ToString());
                    if (rowSpan > 1)
                        tc.SetAttributeValue("rowSpan", rowSpan.ToString());
                }
                else
                {
                    // Spanned cells: set merge flags
                    if (c > minCol)
                        tc.SetAttributeValue("hMerge", "1");
                    if (r > minRow)
                        tc.SetAttributeValue("vMerge", "1");
                }
            }
        }

        _slidePart?.Save();

        // Return the origin cell
        var rows = Rows;
        return rows[minRow][minCol];
    }

    /// <inheritdoc />
    public void SetTextFormat(IBasePortionFormat source)
    {
        var cells = CollectAllCells();
        TextFormatHelper.ApplyTextFormat(cells, source, _slidePart);
    }

    /// <inheritdoc />
    public void SetTextFormat(IParagraphFormat source)
    {
        var cells = CollectAllCells();
        TextFormatHelper.ApplyTextFormat(cells, source, _slidePart);
    }

    /// <inheritdoc />
    public void SetTextFormat(ITextFrameFormat source)
    {
        var cells = CollectAllCells();
        TextFormatHelper.ApplyTextFormat(cells, source, _slidePart);
    }

    // ────── Private helpers ──────

    private bool GetBoolAttr(string attrName)
    {
        return _tblPr?.Attribute(attrName)?.Value == "1";
    }

    private void SetBoolAttr(string attrName, bool value)
    {
        var tblPr = EnsureTblPr();
        if (value)
            tblPr.SetAttributeValue(attrName, "1");
        else
            tblPr.Attribute(attrName)?.Remove();
        _slidePart?.Save();
    }

    private XElement EnsureTblPr()
    {
        if (_tblPr is not null)
            return _tblPr;

        if (_tbl is not null)
        {
            _tblPr = new XElement(ANs + "tblPr");
            _tbl.AddFirst(_tblPr);
            return _tblPr;
        }

        // Detached table (no XML tree): create a standalone tblPr for in-memory use.
        _tbl = new XElement(ANs + "tbl");
        _tblPr = new XElement(ANs + "tblPr");
        _tbl.Add(_tblPr);
        _tblGrid = new XElement(ANs + "tblGrid");
        _tbl.Add(_tblGrid);
        return _tblPr;
    }

    private static XElement? FindStyleElement(XElement tblPr)
    {
        return tblPr.Element(ANs + "tblStyleId")
            ?? tblPr.Element(ANs + "tblStyle");
    }

    private static string ReadStyleGuid(XElement? styleEl)
    {
        if (styleEl is null)
            return string.Empty;
        // <a:tblStyleId>GUID</a:tblStyleId>
        if (!string.IsNullOrWhiteSpace(styleEl.Value))
            return styleEl.Value.Trim();
        // Legacy: <a:tblStyle val="GUID"/>
        return styleEl.Attribute("val")?.Value ?? string.Empty;
    }

    private List<ICell> CollectAllCells()
    {
        var cells = new List<ICell>();
        var rows = Rows;
        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (var c = 0; c < row.Count; c++)
                cells.Add(row[c]);
        }
        return cells;
    }
}
